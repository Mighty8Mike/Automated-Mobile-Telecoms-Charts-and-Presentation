import os

from ITU_Utilities import load_and_prepare_data, prepare_slides, print_slides, delete, CHARTS_PATH, SLIDES_PATH
from Create_Charts import create, select_image_files, read_entry


# Define paths
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_PATH, "formatted_for_sbrn.xlsx")
CHARTS_PATH = os.path.join(BASE_PATH, 'Charts')
SLIDES_PATH = os.path.join(BASE_PATH, 'Slides')
PRESENTATIONS_PATH = os.path.join(BASE_PATH, 'Presentations')





# Ensure necessary directories exist
def ensure_dir(path):
    if not os.path.exists(path):
        os.mkdir(path)

# Print the main menu
def print_menu():
    print("\nMenu:")
    print("0. Load dataframe")
    print("1. Create new charts")
    print("2. Read charts")
    print("3. Prepare slides")
    print("4. Print slides")
    print("5. Delete charts and slides")
    print("6. Exit")

# Main control function
def main():
    ensure_dir(PRESENTATIONS_PATH)
    ensure_dir(CHARTS_PATH)
    ensure_dir(SLIDES_PATH)

    df = None  # Will hold the loaded dataframe

    while True:
        print_menu()
        choice = input("Choose action: ").strip()

        if choice == "0":
            try:
                df = load_and_prepare_data(EXCEL_PATH)
                print("\u2705 DataFrame loaded and formatted.")
            except Exception as e:
                print(f"\u274C Failed to load DataFrame: {e}")

        elif choice == "1":
            if df is None:
                print("\u26A0\uFE0F Please load the dataframe first (option 0).")
            else:
                try:
                    create(df, CHARTS_PATH)
                except Exception as e:
                    print(f"\u274C Chart creation failed: {e}")

        elif choice == "2":
            try:
                paths = select_image_files(CHARTS_PATH)
                if paths:
                    read_entry(paths)  # Show images selected the first time
                else:
                    print("\u274C No files selected.")
            except Exception as e:
                print(f"\u274C Failed to read charts: {e}")


        elif choice == "3":
            try:
                prepare_slides(SLIDES_PATH, CHARTS_PATH)
            except Exception as e:
                print(f"\u274C Prepare slides failed: {e}")

        elif choice == "4":
            try:
                print_slides()
            except Exception as e:
                print(f"\u274C Print failed: {e}")

        elif choice == "5":
            try:
                delete()
            except Exception as e:
                print(f"\u274C Delete failed: {e}")

        elif choice == "6" or choice.lower() == "exit":
            print("\U0001F44B Finished, bye!")
            break

        else:
            print("\u2757 Not an option! Try again.")

# Guard main execution
if __name__ == "__main__":
    main()

