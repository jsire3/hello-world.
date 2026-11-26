import threading
import time

from conftest import FakeOffice

from doc2word.worker import ConversionQueue, State

FINISHED = {State.DONE, State.WARNING, State.FAILED, State.CANCELLED}


def _run(jobs_src, office, timeout_s=30, out_dir=None):
    results = {}
    done = threading.Event()

    def on_update(job):
        if job.state in FINISHED:
            results[job.src.name] = (job.state, job.message)
            if len(results) == len(jobs_src):
                done.set()

    q = ConversionQueue(on_update, session_factory=lambda: office, timeout_s=timeout_s)
    for s in jobs_src:
        q.submit(s, out_dir)
    assert done.wait(20), results
    q.shutdown()
    return results


def test_queue_mixed_outcomes(tmp_path, digital_pdf, scanned_pdf, png_file):
    locked = tmp_path / "locked.txt"
    locked.write_text("x")
    office = FakeOffice(fail_on="input.txt")
    res = _run([digital_pdf, scanned_pdf, png_file, locked], office)
    assert res["digital report.pdf"][0] == State.DONE
    assert res["scan.pdf"][0] == State.WARNING
    assert res["photo.png"][0] == State.WARNING
    assert res["locked.txt"] == (State.FAILED, "The file is password protected.")
    assert office.resets == 1   # Office restarted after the failure
    assert office.closed        # and closed on shutdown


def test_watchdog_kills_hung_office(digital_pdf, png_file):
    office = FakeOffice(hang_s=10)
    start = time.monotonic()
    res = _run([digital_pdf, png_file], office, timeout_s=1)
    assert res["digital report.pdf"][0] == State.FAILED
    assert "longer than 1 seconds" in res["digital report.pdf"][1]
    assert res["photo.png"][0] == State.WARNING  # queue carried on
    assert time.monotonic() - start < 8


def test_gui_end_to_end(qapp_and_window, digital_pdf, png_file, tmp_path):
    app, win = qapp_and_window
    (tmp_path / "ignored.zip").write_bytes(b"PK")
    assert win.add_paths([tmp_path]) == 2       # folder expanded, .zip skipped
    assert win.add_paths([digital_pdf]) == 0    # no duplicates
    assert win.b_convert.isEnabled()
    win.open_when_done.setChecked(False)
    win.start()

    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        app.processEvents()
        statuses = [win.tree.topLevelItem(i).text(2) for i in range(win.tree.topLevelItemCount())]
        if all(s in (State.DONE.value, State.WARNING.value) for s in statuses):
            break
        time.sleep(0.05)
    assert sorted(statuses) == sorted([State.DONE.value, State.WARNING.value]), statuses
    assert (tmp_path / "digital report.docx").is_file()
    assert (tmp_path / "photo.docx").is_file()
    assert win.b_clear.isEnabled()
    win._clear_finished()
    assert win.tree.topLevelItemCount() == 0
