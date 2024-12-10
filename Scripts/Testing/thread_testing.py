import threading
import queue
import time
import tkinter as tk
from tkinter import ttk

# Dummy fuzzy matching function
def fuzzy_match(data, progress_queue):
    result = []
    for i, item in enumerate(data):
        # Simulate matching operation
        time.sleep(0.1)
        result.append(f"Matched: {item}")
        # Send progress update
        progress_queue.put((i + 1, len(data)))
    # Send the final result to the main thread
    progress_queue.put(("result", result))

# Main GUI application
def run_app():
    def start_matching():
        # Clear previous results
        output_text.delete("1.0", tk.END)
        progress_bar["value"] = 0

        # Input data for matching
        data = [f"Item {i}" for i in range(10)]

        # Start the worker thread
        worker_thread = threading.Thread(
            target=fuzzy_match, args=(data, progress_queue), daemon=True
        )
        worker_thread.start()

        # Start checking for progress updates
        check_progress()

    def check_progress():
        try:
            while True:
                update = progress_queue.get_nowait()
                if update[0] == "result":
                    # Matching is complete; display the result
                    result = update[1]
                    output_text.insert("1.0", "\n".join(result))
                    return
                else:
                    # Update progress bar
                    completed, total = update
                    progress_bar["value"] = (completed / total) * 100
        except queue.Empty:
            # No updates in the queue; continue checking
            root.after(100, check_progress)

    # Create GUI
    root = tk.Tk()
    root.title("Fuzzy Matching Tool")

    progress_bar = ttk.Progressbar(root, maximum=100)
    progress_bar.pack(fill=tk.X, padx=10, pady=10)

    start_button = ttk.Button(root, text="Start Matching", command=start_matching)
    start_button.pack(pady=10)

    output_text = tk.Text(root, height=10, width=50)
    output_text.pack(padx=10, pady=10)

    root.mainloop()

# Queue for thread communication
progress_queue = queue.Queue()

if __name__ == "__main__":
    run_app()
