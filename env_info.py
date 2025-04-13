import sys
import subprocess

def print_python_info():
    print("Python Version:", sys.version)
    print("Python Executable:", sys.executable)

def print_installed_packages():
    try:
        # Using pip freeze for detailed versions
        result = subprocess.run(
            ["pip", "freeze"],
            capture_output=True, text=True, check=True
        )
        print("\nInstalled Packages:")
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print("Error retrieving installed packages:", e)

if __name__ == "__main__":
    print_python_info()
    print_installed_packages()