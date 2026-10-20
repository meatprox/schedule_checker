"""画面を起動して、サンプルの読み込み、拡大、編集、保存までを確認する。"""

import tempfile
import unittest
from datetime import timedelta
from pathlib import Path

import tkinter as tk

from app import Drag, ScheduleApp, enable_dpi
from model import parse_workbook


class AppTests(unittest.TestCase):
    def test_sample_zoom_edit_and_save(self):
        enable_dpi()
        root = tk.Tk()
        root.geometry("1100x720")
        try:
            app = ScheduleApp(root)
            app.load_sample()
            root.update_idletasks()
            root.update()
            self.assertEqual(len(app.tasks), 10)
            self.assertGreater(len(app.point_hits), 0)
            self.assertGreater(app.timeline_width, 1000)

            before = app.px_per_hour
            app.zoom_by(200, 120)
            self.assertGreater(app.px_per_hour, before)
            app.zoom_by(200, -120)
            self.assertAlmostEqual(app.px_per_hour, before, delta=0.2)

            morning = next(task for task in app.tasks if task.name == "朝会")
            origin = [point.time for point in morning.points]
            morning.segments[0].style = "dashed"
            for point in morning.points:
                point.time += timedelta(minutes=30)
            app.update_status()
            self.assertNotIn("disabled", app.reset_button.state())

            with tempfile.TemporaryDirectory() as folder:
                path = Path(folder) / "sample_編集.xlsx"
                app.save_to(path)
                parsed = parse_workbook(path)
                restored = next(task for task in parsed.tasks if task.points[0].time.hour == 9 and task.points[0].time.minute == 30)
                self.assertEqual(restored.segments[0].style, "dashed")
                self.assertEqual(restored.points[0].time, origin[0] + timedelta(minutes=30))

            app.restore()
            morning = next(task for task in app.tasks if task.name == "朝会")
            self.assertEqual([point.time for point in morning.points], origin)
            self.assertEqual(morning.segments[0].style, "solid")
            self.assertIn("disabled", app.reset_button.state())

            point = morning.points[1]
            slot = morning.slot
            target = point.time + timedelta(minutes=8)
            app.drag = Drag("point", task_id=morning.id, point_id=point.id, origin_time=point.time, slot=slot)
            app.move_point(app.drag, self._event(app, app.time_to_x(target), 40))
            self.assertEqual(morning.slot, slot)
            self.assertEqual(point.time, target)

            origins = [item.time for item in morning.points]
            start = app.time_to_x(morning.points[0].time)
            app.drag = Drag(
                "task",
                task_id=morning.id,
                origin_slot=morning.slot,
                origin_times=list(origins),
                slot=morning.slot,
                start_canvas_x=start,
            )
            slot_y = app.content_top() + 4 * app.slot_h + 12
            app.move_task(app.drag, self._event(app, start + app.px_per_hour, slot_y))
            self.assertEqual(morning.slot, 4)
            self.assertEqual(morning.points[0].time, origins[0] + timedelta(hours=1))
            self.assertEqual(morning.points[1].time - morning.points[0].time, origins[1] - origins[0])
        finally:
            root.destroy()

    def _event(self, app, canvas_x, canvas_y):
        class Event:
            pass

        event = Event()
        event.x_root = app.canvas.winfo_rootx() + (canvas_x - app.canvas.canvasx(0))
        event.y_root = app.canvas.winfo_rooty() + canvas_y
        return event


if __name__ == "__main__":
    unittest.main()
