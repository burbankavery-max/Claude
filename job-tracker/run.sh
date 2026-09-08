#!/bin/bash
# Daily refresh. Logs to run.log; keeps the last 500 lines.
cd "$(dirname "$0")" || exit 1
./.venv/bin/python -u job_tracker.py run >> run.log 2>&1
tail -n 500 run.log > run.log.tmp && mv run.log.tmp run.log
