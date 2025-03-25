 GUI-based personal task manager built with Python.  
It combines **ABCD prioritization**, **Pomodoro technique**, **urgency detection**, and a live **dashboard** — all offline, all in one file.

---

## 🚀 Features

### ✅ Task Management
- Import tasks from `tasks.xlsx`
- ABCD Priority matrix (Urgent/Important logic)
- Smart urgency detection based on:
  - Due date
  - Effort estimate
  - Current progress

### ⏱ Pomodoro Cycles
- 25-minute work sessions
- 5-minute short breaks
- 15-minute long break every 4 Pomodoros
- Auto-switching phases and reminders

### 📊 Dashboard Tab
- Pomodoros completed per day (weekly line chart)
- Tracks your focus and flow over time
- Powered by `pomodoro_log.csv` (auto-generated)

### 📈 Progress Tracking
- Updates Excel file with progress (%) after each Pomodoro
- Visual schedule with urgency flags
- Live data refresh after each timer

---

## 🧪 How It Works

1. You define tasks in `tasks.xlsx` using this format:
 it uses the ABCD  task management method all estimates are in units of pomodory

2. Launch the app with:
```bash
python tasktracker.py
