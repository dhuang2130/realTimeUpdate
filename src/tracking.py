import subprocess
import signal
import time
import os
import platform

# Get the directory of the current script
current_dir = os.path.dirname(os.path.abspath(__file__))

# Construct paths to the scripts
script1 = os.path.join(current_dir, 'realTime.py')
script2 = os.path.join(current_dir, 'realTimeManufacture.py')

# Function to run a script and log its execution
def run_script(script_path):
    print(f"Starting {script_path}")
    if platform.system() == 'Windows':
        process = subprocess.Popen(['python', script_path], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
    else:
        process = subprocess.Popen(['python3', script_path])
    return process

# Function to stop the subprocesses
def stop_processes(processes):
    for process in processes:
        print(f"Stopping process {process.pid}")
        process.terminate()
        process.wait()

# Run both scripts as subprocesses with logging
process1 = run_script(script1)
process2 = run_script(script2)

# Store the processes in a list for easy management
processes = [process1, process2]

try:
    # Keep the main script running to allow for real-time updates
    print("Scripts are running. Press Ctrl+C to stop.")
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("Received stop signal. Terminating subprocesses...")
    stop_processes(processes)

print(f"Finished running {script1}")
print(f"Finished running {script2}")
