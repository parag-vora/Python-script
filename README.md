This Python script processes Wireshark-generated .csv files to analyze and calculate various communication latencies in an industrial Ethernet network.
The script detects specific payload patterns such as "FFFF_FFFF" and "0000_0000" from the Data column and calculates latencies including Bus-to-Bus, Comm-to-App, App-to-Comm, and Comm-to-Bus.
It processes data row-by-row, isolating critical payload segments based on source-destination pairs (e.g., Siemens to Texas) and internal communication. 
All results are saved to an Excel file for detailed analysis, providing insights into network performance. Paths for input and output files are configurable for user flexibility.

## Key Steps
1. File Path Setup:
   * Set paths to both the .csv file (input) and the Excel file (output) where results will be saved.

2. Parameter Initialization:
    * Begin by setting all parameters (e.g., latency, cumulative_time, etc.) to zero to ensure accurate calculations.
  
3. CSV Row-by-Row Processing:
    * Open the CSV file and read it row-by-row.
    * Each row contains data in columns: packet number, time since previous packet, source, destination, payload, etc.

4. Payload Selection Based on Source and Destination:
    * For packets from siemens (source) to texas (destination), consider the payload[16:24] data. Similarly, use payload[14:22] and internal_info[25:35] for other specific packet types.
    * This helps isolate the relevant data within the payload for further calculations.

5. Detecting Pattern Conditions:
    * For detecting latencies based on the packet sequence:
       - Example Condition: If the last packet source is texas, destination is siemens, and the payload is "00000000", then the next packet should ideally be source=siemens, destination=texas, and payload="ffffffff".
       - When this condition is satisfied, freeze the packet time (reference_packet_time = time_since_previous_packet).
       - Calculate Bus-to-Bus latency by summing packets (awaiting_accumulation).

 6. Latency Calculation:
    * Continue adding time_since_previous_packet until the desired packet sequence occurs.
    * Compute latency = cumulative_time - reference_packet_time.

 7. Internal Latency Calculations:
     * Comm-App Latency: This is calculated by cumulative_internal_time - Bus_Comm_latency when internal_info payload equals "0x00000004".
     * App-Comm Latency: Set as cumulative_internal_time when internal_info payload equals "0x00000002".
     * Similar calculations apply for Bus-Comm and Comm-Bus latencies:
        - Bus_Comm_latency = time_since_previous_packet.
        - Comm_Bus_latency = cumulative_time - reference_packet_time - Bus_Comm_latency - Comm_App_latency - App_Comm_latency   

 8. Write Results to Excel:
     * Use the function write_to_excel() to save each latency type (latencies, Comm_App_latencies, App_Comm_latencies, Bus_Comm_latencies, and Comm_Bus_latencies) into the designated Excel file.
