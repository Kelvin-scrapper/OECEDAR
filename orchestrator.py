import subprocess
import sys
import os

MAP_RETRY_TRIGGERS = [
    "Could not find 'Time period' row",
    "No data was parsed",
]

MAX_RETRIES = 3

def run_script(script_name):
    """Runs a python script, prints its output, and returns (exit_code, full_output)."""
    print(f'--- Running {script_name} ---')
    output_lines = []
    try:
        env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
        process = subprocess.Popen(
            [sys.executable, script_name],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', env=env,
        )
        while True:
            line = process.stdout.readline()
            if line == '' and process.poll() is not None:
                break
            if line:
                print(line.strip())
                output_lines.append(line)
        rc = process.poll()
        if rc == 0:
            print(f'--- {script_name} finished successfully ---\n')
        else:
            print(f'--- {script_name} failed with return code {rc} ---\n')
        return rc, ''.join(output_lines)
    except FileNotFoundError:
        print(f"Error: {script_name} not found.")
        return 1, ''

def main():
    """Orchestrates the execution of main.py and map.py, with auto-retry on bad downloads."""
    print('--- Starting Orchestrator ---')

    for attempt in range(1, MAX_RETRIES + 1):
        if attempt > 1:
            print(f'\n--- RETRY {attempt}/{MAX_RETRIES}: Re-running full pipeline ---\n')

        # Run main.py to download the data
        rc, _ = run_script('main.py')
        if rc != 0:
            print('main.py failed. Aborting.')
            return

        # Run map.py to process the data
        rc, map_output = run_script('map.py')

        # Check if map.py hit a bad-download error that warrants a retry
        needs_retry = any(trigger in map_output for trigger in MAP_RETRY_TRIGGERS)

        if rc == 0 and not needs_retry:
            print('--- Orchestrator finished successfully ---')
            return

        if needs_retry and attempt < MAX_RETRIES:
            print(f'[RETRY] map.py could not parse the downloaded file — re-downloading...')
            continue

        # Exhausted retries or non-retryable failure
        if needs_retry:
            print(f'[FAILURE] map.py parsing failed after {MAX_RETRIES} attempts.')
        else:
            print('map.py failed.')
        return

    print('--- Orchestrator finished ---')

if __name__ == '__main__':
    main()
