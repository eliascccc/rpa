# LocalRPA Orchestrator

A lightweight local orchestrator for email- and data-driven automation, delegating execution to external screen-based RPA tools.

---

## Overview

This project is a lightweight local RPA orchestrator written in Python.

It is designed as a simple alternative to heavy enterprise orchestrators,
focusing on clarity, ease of modification, and running on a single machine.

It does NOT replace RPA tools.

Instead, it orchestrates them:
you still need a real RPA tool (Power Automate, UiPath Studio, Blue Prism, etc.)
to perform screen-based automation.

---


## Typical examples

A typical email-driven flow is that a user sends an email to the robot asking it to perform a task. The orchestrator reads the request, validates the input, prepares the required payload, and writes the handover state. The front-end RPA tool then picks up the job and performs the UI actions.

Another main flow is data-driven. In that case, the orchestrator itself discovers work by polling a query or another data source. When it finds a valid case, it prepares the required values and signals the front-end RPA to execute the task.

---

## Key Idea

This project separates **orchestration** from **UI automation**:

* The **back-end (this project)** handles:

  * job intake (email / data)
  * validation
  * decision logic
  * audit logging
  * system control

* The **front-end RPA tool** handles:

  * clicks
  * keyboard input
  * ERP/UI interaction

They communicate through a file-based IPC mechanism (`handover.json`).

---

## Architecture

<img width="1156" height="1921" alt="workflow" src="https://github.com/user-attachments/assets/aadd3936-756c-42f3-9d8a-333dd48fbbf0" />

The diagram shows how:

* Python (backend) and RPA (front-end) run independently
* Both operate in their own loops
* State is synchronized via handover.txt
* Failures transition the system into safestop


## Features

* Email-driven job processing (own inbox)
* Shared inbox support (extensible)
* Data-driven jobs (ERP/query simulation)
* File-based IPC (`handover.json`)
* SQLite audit logging (`job_audit.db`)
* Crash-safe mode (`safestop`)
* Manual reboot mechanism (`reboot.flag`)
* Network-aware execution
* Built-in screen recording (ffmpeg)
* Works without admin rights
* Easy to share the full runtime with an AI assistant
* Runs on both Windows and Linux

---

## Design Principles

* **Simplicity over scalability**
* **Local-first execution**
* **Fail fast and visibly**
* **No hidden state (file + DB only)**
* **Deterministic job lifecycle**
* **Cheap to deploy and operate**

---

## Job Sources

The system supports multiple job producers:

* Personal inbox (`personal_inbox`)
* Shared mailbox (planned/partial)
* Scheduled jobs (ERP/data queries)

All sources produce standardized **candidates**, processed through a unified flow.

---

## Job Lifecycle

Jobs are tracked in SQLite (`audit.db`) with clear states:

* `REJECTED` – invalid request / no access
* `QUEUED` – waiting for RPA
* `RUNNING` – RPA executing
* `VERIFYING` – post-check (Python)
* `DONE` – success
* `FAIL` – error or verification failure

---

## Handover concept

The system uses a file-based "handover" mechanism to transfer control between the Python orchestrator and the external RPA.

The handover file represents both:
- the current system state
- and the payload required for the next execution step

---

## Email Pipeline

* Access controlled via `friends.xlsx`
* Only allowed users can trigger jobs
* Job permissions defined per user
* Unknown senders → silently deleted
* Automatic replies:

  * `DONE`
  * `FAIL`
  * "lifesign" (once per day)

---

## Running the Project

### Requirements

* Python 3.14
* `openpyxl`
* `ffmpeg` (optional, for screen recording)

---

### Start

```bash
python main.py
```

---

### Test setup

Use included dev tools:

* `fake_jobs_generator.py`
* `frontend_rpa_simulator.py`

---

## Project Structure (simplified)

```
main.py
own_inbox/
shared_inbox/
handover.txt
audit.db
friends.xlsx
recordings/
```

---

## Intended Use Case

* Small internal automation (5–10 users)
* No dedicated RPA infrastructure
* No admin rights required
* Cheap “extra laptop” deployment
* Pilot / proof-of-concept automation

---

## Limitations

* Not designed for large-scale orchestration
* No distributed execution
* File-based IPC only
* Minimal error recovery (by design)

---

## Why not just use Robot Framework?

| Robot Framework                    | LocalRPA Orchestrator          |
| ---------------------------------- | ------------------------------ |
| Scalable test/automation framework | Lightweight local orchestrator |
| Requires setup/integration         | Runs on a single machine       |
| Code-driven automation             | Combines email + RPA tools     |
| General-purpose                    | Business-triggered workflows   |

---

## Philosophy

> This is not a full RPA platform.
> It is the simplest possible glue between:
>
> * business triggers (email/data)
> * and UI automation (RPA tools)

---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

readme written by AI
