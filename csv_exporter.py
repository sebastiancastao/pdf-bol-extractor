import os
import gc
import glob
import pandas as pd
from config import OUTPUT_CSV_NAME

class CSVExporter:
    def __init__(self, session_dir):
        """Initialize the CSV exporter with a session directory."""
        self.session_dir = session_dir

    def combine_to_csv(self):
        """Combine all CSV files in the session directory into one."""
        try:
            # Get all CSV files in the session directory except the output file
            csv_files = [f for f in glob.glob(os.path.join(self.session_dir, "*.csv"))
                        if os.path.basename(f) != OUTPUT_CSV_NAME]

            if not csv_files:
                print("No CSV files found to combine")
                return False

            print(f"Found {len(csv_files)} CSV files to combine")

            # Process files in chunks to conserve memory
            chunk_size = 5
            output_path = os.path.join(self.session_dir, OUTPUT_CSV_NAME)
            first_file = True

            for i in range(0, len(csv_files), chunk_size):
                chunk = csv_files[i:i + chunk_size]
                print(f"Processing chunk {i//chunk_size + 1} of {(len(csv_files) + chunk_size - 1)//chunk_size}")
                
                # Read and combine chunk of CSV files
                dfs = []
                files_to_remove = []
                for file in chunk:
                    try:
                        # Read CSV in chunks
                        for df_chunk in pd.read_csv(file, chunksize=1000, dtype=str):
                            dfs.append(df_chunk)
                        files_to_remove.append(file)  # Mark for removal after reading
                    except Exception as e:
                        print(f"Error processing {file}: {str(e)}")
                        continue

                if not dfs:
                    continue

                # Combine chunks and write to output
                chunk_df = pd.concat(dfs, ignore_index=True)
                
                if first_file:
                    # Write with header for first chunk
                    chunk_df.to_csv(output_path, index=False, mode='w')
                    first_file = False
                else:
                    # Append without header for subsequent chunks
                    chunk_df.to_csv(output_path, index=False, mode='a', header=False)

                # **ROBUST CLEANUP**: Remove individual CSV files after successful combining
                for file in files_to_remove:
                    max_attempts = 3
                    for attempt in range(max_attempts):
                        try:
                            if os.path.exists(file):
                                os.remove(file)
                                print(f"✅ Removed individual CSV: {os.path.basename(file)}")
                                break
                        except Exception as e:
                            if attempt < max_attempts - 1:
                                print(f"⚠️ Attempt {attempt + 1} failed to remove {file}: {str(e)}")
                                import time
                                time.sleep(0.1)  # Brief wait before retry
                            else:
                                print(f"❌ FAILED to remove {file} after {max_attempts} attempts: {str(e)}")
                                print(f"❌ This may cause contamination detection in future workflows")

                # Clear memory
                del dfs
                del chunk_df
                gc.collect()

            print(f"Successfully combined files into {OUTPUT_CSV_NAME}")
            return True

        except Exception as e:
            print(f"Error combining CSV files: {str(e)}")
            return False

if __name__ == "__main__":
    exporter = CSVExporter(".")
    exporter.combine_to_csv()
