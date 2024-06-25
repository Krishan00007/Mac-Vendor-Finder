# import openpyxl, pandas and mac_vendor_lookup library.
from mac_vendor_lookup import MacLookup

maclook = MacLookup()

def mac_vendor_finder(mac):
    result = str("")
    try:
        result = maclook.lookup(mac) 
    except:
        result = str("none")

    print(result)

def main():

    # call the mac_vendor_finder function in recursive loop
    check = 1
    while (check == 1):
        mac = str(input("Enter the Mac(format:ff:ff:ff:ff:00): "))
        mac_vendor_finder(mac)
        
        check = int(input("check again:"))
        

if __name__ == "__main__":
    main()
