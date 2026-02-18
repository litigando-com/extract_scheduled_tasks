# schtasks /create /sc once /tn "prueba_sencilla" /tr "notepad.exe" /st 18:00
import subprocess

def read_file():
    with open("tasks.txt", "w") as file:
        try:
            result = subprocess.run(
                ["schtasks", "/query"],
                    capture_output=True,
                    text=True
                )

            lines = result.stdout.split("\n")
            
            print("Reading file...\n")

            for line in lines:
                # file.write(f"{line}\n")
                print(f"{line}")

            print("Processes inserted into tasks.txt!\n")

        except Exception as e:
            print(f"error with the file: {e}")

def main():
    print("=============================")
    print("\n== Read tasks from Windows ==\n")
    print("=============================")

    read_file();

if __name__ == "__main__":
    main()
