import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import rcParams
from colorama import Fore, Style, init

# Initialize colorama
init(autoreset=True)

# Set font for Chinese characters
rcParams['font.sans-serif'] = ['SimSun']  # Set to SimSun
rcParams['axes.unicode_minus'] = False   # Fix the issue with displaying minus signs

def moving_average(data, window_size):
    """Calculate moving average for data smoothing."""
    return data.rolling(window=window_size, center=True).mean()

def analyze_and_plot(file_path, output_plot_path):
    print(Fore.CYAN + "Analyzing the file and detecting events...")

    # Read Excel file
    data = pd.read_excel(file_path, engine='openpyxl')

    # Clean column names by removing spaces and converting to lowercase
    data.columns = data.columns.str.strip().str.lower()

    # Check required columns
    required_columns = [
        'voltage', 'voltage_0', 
        'voltage (dc voltage)', 'voltage_0 (dc voltage)', 
        'voltage (positive peak)', 'voltage_0 (negative peak)',
        'voltage_0 (positive peak)'
    ]
    if not all(col in data.columns for col in required_columns):
        raise ValueError(Fore.RED + f"Excel file is missing required columns: {required_columns}")

    # Assume each data point has a time interval of 0.002 seconds
    time_interval = 0.002
    data['time'] = [i * time_interval for i in range(len(data))]

    # Smooth the data
    window_size = 15
    data['voltage_smooth'] = moving_average(data['voltage'], window_size)
    data['voltage_0_smooth'] = moving_average(data['voltage_0'], window_size)

    # Read baseline and threshold values
    vertical_baseline = data['voltage (dc voltage)'].iloc[0]
    horizontal_baseline = data['voltage_0 (dc voltage)'].iloc[0]
    t_vertical_threshold = data['voltage (positive peak)'].iloc[0]
    if data['voltage_0 (negative peak)'].iloc[0] < data['voltage_0 (positive peak)'].iloc[0]:
        t_horizontal_threshold = data['voltage_0 (negative peak)'].iloc[0]
    else:
        t_horizontal_threshold = data['voltage_0 (positive peak)'].iloc[0]

    # Set parameters
    threshold_duration = 0.4  # Duration threshold to distinguish blink from action potential
    cooldown_time = 0.5  # Cooldown time after action potential (seconds)
    time_advance = 0.05  # Advance end_time by 0.05 seconds
    vertical_threshold = t_vertical_threshold
    horizontal_threshold = t_horizontal_threshold

    # Store events
    events = []

    # Detect blink and action potential events
    start_time = None
    is_in_event = False
    last_event_end_time = None  # Record the end time of the last event
    triggering_type = None

    for i in range(len(data)):
        current_time = data['time'].iloc[i]
        current_vertical = data['voltage_smooth'].iloc[i]
        current_horizontal = data['voltage_0_smooth'].iloc[i]

        # Check cooldown period: skip detection if within cooldown period after the last event
        if last_event_end_time and current_time < last_event_end_time + cooldown_time:
            continue

        # Check if the vertical or horizontal voltage exceeds the threshold
        if abs(current_vertical - vertical_baseline) >= vertical_threshold or abs(current_horizontal - horizontal_baseline) >= 3*horizontal_threshold:
            if not is_in_event:
                # Event starts
                is_in_event = True
                start_time = current_time
                triggering_type = "horizontal" if abs(current_horizontal - horizontal_baseline) >= 3*horizontal_threshold else "vertical"
        elif is_in_event:
            # Check if back to baseline range
            if triggering_type == "horizontal" and abs(current_horizontal - horizontal_baseline) <= 3*horizontal_threshold:
                end_time = current_time - time_advance
                duration = end_time - start_time

                # Find the corresponding horizontal value at end_time
                end_horizontal = data['voltage_0_smooth'].iloc[(data['time'] - end_time).abs().argmin()]

                event_type = "Action potential"
                direction = "Horizontal: " + ("Looking left" if end_horizontal > horizontal_baseline else "Looking right")

                # Record event
                events.append((start_time, end_time, event_type, direction))
                last_event_end_time = end_time

                # Reset event status
                is_in_event = False
                start_time = None
                triggering_type = None

            elif triggering_type == "vertical" and abs(current_vertical - vertical_baseline) <= 3*vertical_threshold:
                end_time = current_time - time_advance
                duration = end_time - start_time

                # Find the corresponding vertical and horizontal values at end_time
                end_vertical = data['voltage_smooth'].iloc[(data['time'] - end_time).abs().argmin()]
                end_horizontal = data['voltage_0_smooth'].iloc[(data['time'] - end_time).abs().argmin()]

                # Determine event type
                if duration <= threshold_duration:
                    event_type = "Blink"
                    direction = None
                else:
                    event_type = "Action potential"
                    direction = "Vertical: " + ("Looking up" if end_vertical > vertical_baseline else "Looking down")
                    if abs(end_horizontal - horizontal_baseline) >= horizontal_threshold:
                        direction += ", Horizontal: " + ("Looking left" if end_horizontal > horizontal_baseline else "Looking right")

                # Record event
                events.append((start_time, end_time, event_type, direction))
                last_event_end_time = end_time

                # Reset event status
                is_in_event = False
                start_time = None
                triggering_type = None

    # Plot the graph
    plt.figure(figsize=(12, 6))
    plt.plot(data['time'], data['voltage_smooth'], label='Vertical Signal (Black)', color='black', linewidth=0.8)
    plt.plot(data['time'], data['voltage_0_smooth'], label='Horizontal Signal (Red)', color='red', linewidth=0.8)

    # Highlight events on the plot
    for start_time, end_time, event, direction in events:
        color = 'yellow' if event == 'Blink' else 'blue'
        plt.axvspan(start_time, end_time, color=color, alpha=0.3, label=event)

    plt.title('Potential Event Detection')
    plt.xlabel('Time (seconds)')
    plt.ylabel('Voltage (V)')
    plt.legend(loc='upper right')
    plt.grid(True)
    plt.savefig(output_plot_path, dpi=300, bbox_inches='tight', facecolor='white')

    # Display event list
    print(Fore.GREEN + "=== Detected Events ===")
    for start_time, end_time, event, direction in events:
        if event == "Blink":
            print(Fore.YELLOW + f"Event: {event} - Start: {start_time:.3f}s, End: {end_time:.3f}s")
        else:
            print(Fore.RED + f"Event: {event} - Start: {start_time:.3f}s, End: {end_time:.3f}s, Direction: {direction}")
    print("\n")

def main():
    default_output_path = None

    while True:
        
        print(Fore.CYAN + "\n--- EOG Event Detection ---")
        print(Fore.MAGENTA + "1. Analyze an Excel file")
        print(Fore.MAGENTA + "2. Set default output path")
        print(Fore.MAGENTA + "3. Exit")
       
        choice = input(Fore.YELLOW + "Enter your choice: ")
        print("\n")
        if choice == "1":
            file_path = input(Fore.CYAN + "Enter the path to the Excel file: ")

            if default_output_path:
                output_plot_path = default_output_path
            else:
                output_plot_path = input(Fore.CYAN + "Enter the output path for the plot: ")

            try:
                analyze_and_plot(file_path, output_plot_path)
            except Exception as e:
                print(Fore.RED + f"Error: {e}")

        elif choice == "2":
            default_output_path = input(Fore.CYAN + "Enter the default output path for plots: ")
            print(Fore.GREEN + f"Default output path set to: {default_output_path}")

        elif choice == "3":
            print(Fore.CYAN + "Exiting the program.")
            break

        else:
            print(Fore.RED + "Invalid choice. Please try again.")

if __name__ == "__main__":
    main()