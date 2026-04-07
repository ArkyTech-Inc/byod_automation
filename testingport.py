"""
 Port Testing Script
 This script tests connectivity to specified ports on the Gmail SMTP server.
 It checks if ports 587, 465, and 25 are open or blocked from your network.
 You can run this script to diagnose email sending issues related to port blocking.
 i faced issues running the automation script because my network was blocking port 587, which is commonly used for SMTP.
 Mainly because i was connected to the NITDA Wifi and it blocked the port. This script helped me identify that issue.
 Just run: python testingport.py
"""
import socket

ports = [587, 465, 25]
host = "smtp.gmail.com"

for port in ports:
    try:
        sock = socket.create_connection((host, port), timeout=5)
        print(f"Port {port}: OPEN ✓")
        sock.close()
    except Exception as e:
        print(f"Port {port}: BLOCKED ✗ ({e})")