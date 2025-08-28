import subprocess
import sys

def run_script(script_name):
    """Runs a python script and prints its output."""
    print(f'--- Running {script_name} ---')
    try:
        process = subprocess.Popen([sys.executable, script_name], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8')
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                print(output.strip())
        rc = process.poll()
        if rc == 0:
            print(f'--- {script_name} finished successfully ---\n')
        else:
            print(f'--- {script_name} failed with return code {rc} ---\n')
        return rc
    except FileNotFoundError:
        print(f"Error: {script_name} not found.")
        return 1

def main():
    """Orchestrates the execution of main.py and map.py."""
    print('--- Starting Orchestrator ---')
    
    # Run main.py to download the data
    if run_script('main.py') != 0:
        print('main.py failed. Aborting.')
        return

    # Run map.py to process the data
    if run_script('map.py') != 0:
        print('map.py failed.')
        return

    print('--- Orchestrator finished ---')

if __name__ == '__main__':
    main()
