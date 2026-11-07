import socket
import re
from datetime import datetime

PORT = 30003

FILE_PATH = r"C:\Users\ADSB-01\Desktop\finalll\final_file.txt"

def extract_ascii(data):
    text = data.decode("ascii", errors="ignore")
    return re.findall(r'[A-Za-z0-9]{3,}', text)

def write_to_file(data, addr):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    hex_data = data.hex()
    ascii_data = extract_ascii(data)

    with open(FILE_PATH, "a", encoding="utf-8") as f:
        f.write("\n" + "="*90 + "\n")
        f.write("PINGSTATION 3 LIVE DATA LOG\n")
        f.write("="*90 + "\n")

        f.write(f"TIME        : {timestamp}\n")
        f.write(f"SOURCE IP   : {addr[0]}\n")
        f.write(f"SOURCE PORT : {addr[1]}\n\n")

        f.write("HEX DATA:\n")
        f.write(hex_data + "\n\n")

        f.write("ASCII DATA:\n")

        if ascii_data:
            f.write(", ".join(ascii_data) + "\n")
        else:
            f.write("None\n")

        f.write("="*90 + "\n")

def main():
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.bind(("0.0.0.0", PORT))
    server.listen(5)

    print(f"🟢 Listening on port {PORT}...\n")

    while True:
        conn, addr = server.accept()
        print(f"✅ Connection from {addr}")

        try:
            while True:
                data = conn.recv(16384)

                if not data:
                    print("⚠️ Connection closed")
                    break

                print(data.hex())

                ascii_data = extract_ascii(data)
                if ascii_data:
                    print("🔤 ASCII:", ascii_data)

                # SAVE TO YOUR FILE PATH
                write_to_file(data, addr)

        except Exception as e:
            print(f"❌ Error: {e}")

        finally:
            conn.close()

if __name__ == "__main__":
    main()
