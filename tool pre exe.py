import os
import sys
import time
import pandas as pd
import subprocess

# Handle different paths for .py and .exe execution
if getattr(sys, 'frozen', False):  # Running as .exe
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Define paths
TOOLS_DIR = os.path.join(script_dir, "TOOLS")  # TOOLS directory containing KAPE and RECmd
DEFAULT_OUTPUT_DIR = r"C:\KAPE_Output"
KAPE_OUTPUT_DIR = os.path.join(DEFAULT_OUTPUT_DIR, "KapeOutput")
RECMD_OUTPUT_DIR = os.path.join(DEFAULT_OUTPUT_DIR, "Parsed_Registry")

# Tool paths
kape_path = os.path.join(TOOLS_DIR, "kape.exe")
recmd_path = os.path.join(TOOLS_DIR, "RECmd.exe")
batch_folder = os.path.join(TOOLS_DIR, "BatchExamples")

# Ensure output directories exist
os.makedirs(KAPE_OUTPUT_DIR, exist_ok=True)
os.makedirs(RECMD_OUTPUT_DIR, exist_ok=True)

# Ask user for command input
command = input("\nEnter command (collect / parse and merge / all): ").strip().lower()

if command in ["collect", "all"]:
    # Step 1: Run KAPE to extract registry hives
    kape_command = [
        kape_path,
        "--tsource", "C:",
        "--target", "RegistryHives",
        "--tdest", KAPE_OUTPUT_DIR,
        "--tflush"
    ]

    print("\n[+] Running KAPE to extract registry hives...")
    try:
        subprocess.run(kape_command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        print("[+] KAPE execution completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error executing KAPE: {e}")
        print(e.stderr)
        sys.exit(1)

if command in ["parse and merge", "all"]:
    # Allow user to specify input and output paths
    kape_input = input("\nEnter path to KAPE output (Press Enter to use default): ").strip() or KAPE_OUTPUT_DIR
    output_path = input("Enter path to save parsed data (Press Enter to use default): ").strip() or RECMD_OUTPUT_DIR
    os.makedirs(output_path, exist_ok=True)

    # Step 2: Locate registry hives
    system32_config_path = os.path.join(kape_input, "C", "Windows", "System32", "config")
    users_folder = os.path.join(kape_input, "C", "Users")

    batch_mappings = {
        "SYSTEM": "DFIRBatch.reb",
        "SOFTWARE": "DFIRBatch.reb",
        "NTUSER.DAT": "DFIRBatch.reb",
        "SAM": "DFIRBatch.reb",
        "SECURITY": "DFIRBatch.reb"
    }

    hive_files = []
    expected_hives = ["SYSTEM", "SAM", "SECURITY", "SOFTWARE"]

    # Step 2.1: Find NTUSER.DAT for all users
    if os.path.exists(users_folder):
        for user in os.listdir(users_folder):
            user_path = os.path.join(users_folder, user, "NTUSER.DAT")
            if os.path.exists(user_path):
                hive_files.append(user_path)
                print(f"[DEBUG] Found NTUSER.DAT for user '{user}' at {user_path}")
    else:
        print("[-] ERROR: Users folder not found in extracted KAPE output.")

    # Step 2.2: Find other registry hives in System32/config
    if os.path.exists(system32_config_path):
        for hive in expected_hives:
            hive_path = os.path.join(system32_config_path, hive)
            if os.path.exists(hive_path):
                hive_files.append(hive_path)
                print(f"[DEBUG] Found {hive} at {hive_path}")
            else:
                print(f"[-] ERROR: {hive} not found in System32\\config.")
    else:
        print("[-] ERROR: System32\\config folder not found.")

    print(f"[+] Found {len(hive_files)} registry hives.")

    # Step 3: Process hives with RECmd
    for hive in hive_files:
        hive_name = os.path.basename(hive)
        batch_file = batch_mappings.get(hive_name)
        batch_path = os.path.join(batch_folder, batch_file) if batch_file else None
        output_csv = os.path.join(output_path, f"{hive_name}.csv")

        if batch_file and os.path.exists(batch_path):
            recmd_command = [
                recmd_path, "-f", hive, "--csv", output_path, "--csvf", f"{hive_name}.csv",
                "--bn", batch_path, "--nl", "--debug"
            ]
            print(f"[+] Using batch file {batch_file} for {hive_name}")
        else:
            recmd_command = [
                recmd_path, "-f", hive, "--csv", output_path, "--csvf", f"{hive_name}.csv",
                "--sa", "--nl", "--debug"
            ]
            print(f"[+] No batch file found. Using --sa for {hive_name}")

        try:
            result = subprocess.run(recmd_command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            print(f"[+] Successfully parsed {hive}. Output saved in {output_csv}")

            if not os.path.exists(output_csv):
                print(f"[-] ERROR: CSV file missing for {hive_name}.")
        except subprocess.CalledProcessError as e:
            print(f"[-] ERROR parsing {hive} with RECmd: {e}")
            print("Error Output:", e.stderr)

    print(f"[+] Parsing complete. Check output at {output_path}")

    # Step 4: Combine all CSV files in Parsed_Registry
    def combine_csv_files(directory, output_filename):
        csv_files = [f for f in os.listdir(directory) if f.endswith(".csv")]
        dfs = []

        for file in csv_files:
            file_path = os.path.join(directory, file)
            try:
                df = pd.read_csv(file_path)
                dfs.append(df)
                print(f"[+] Successfully read {file}")
            except Exception as e:
                print(f"[-] ERROR reading {file}: {e}")

        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            combined_df.to_csv(output_filename, index=False)
            print(f"\n[+] All CSV files have been combined into '{output_filename}'")
        else:
            print("[-] No valid CSV files found to combine.")

    combined_csv_filename = os.path.join(output_path, "combined_registry_data.csv")
    combine_csv_files(output_path, combined_csv_filename)

print("\n[+] Process complete.")
