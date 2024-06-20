import subprocess
import signal
import time

script1 = './src/realTime.py'
script2 = './src/realTimeManufacture.py'

def run_script(script_path):
    print(f"Starting {script_path}")
    process = subprocess.Popen(['python', script_path])
    return process

def stop_processes(processes):
    for process in processes:
        print(f"Stopping process {process.pid}")
        process.send_signal(signal.SIGINT)
        process.wait()

process1 = run_script(script1)
process2 = run_script(script2)

processes = [process1, process2]

try:
    print("Scripts are running. Press Ctrl+C to stop.")
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("Received stop signal. Terminating subprocesses...")
    stop_processes(processes)

print(f"Finished running {script1}")
print(f"Finished running {script2}")