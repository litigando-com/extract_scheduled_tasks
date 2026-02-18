import subprocess

def main():
    with open("tasks.txt", "w") as file:
        try:
            result = subprocess.run(
                ["schtasks", "/query"],
                    capture_output=True,
                    text=True,
                    errors=None
                )

            lines = result.stdout.split("\n")
            
            for line in lines:
                file.write(f"{line}\n")
                # print(f"{line}\n")


        except Exception as e:
            print(f"error {e}")

if __name__ == "__main__":
    main()
