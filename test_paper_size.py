import win32print
import win32con


def set_paper_size(printer_name, size_const):
    try:
        # Open with ADMIN access usually needed for SetPrinter
        # But for defaults maybe not? Let's try standard access first, if fail then we know.
        # Actually SetPrinter needs PRINTER_ACCESS_ADMINISTER for level 2 usually.
        # Check defaults permissions.
        
        PRINTER_ALL_ACCESS = 0xF000C
        hPrinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
        try:
            pInfo = win32print.GetPrinter(hPrinter, 2)
            
            print(f"Old PaperSize: {pInfo['pDevMode'].PaperSize}")
            
            # 9 = A4, 11 = A5
            if pInfo['pDevMode'].PaperSize == size_const:
                print("Size already match.")
                return True
                
            pInfo['pDevMode'].PaperSize = size_const
            pInfo['pDevMode'].Fields |= win32con.DM_PAPERSIZE
            
            # Setup for SetPrinter
            # We must pass the FULL dictionary back
            win32print.SetPrinter(hPrinter, 2, pInfo, 0)
            print("SetPrinter success.")
            return True
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        print(f"Error: {e}")
        return False


if __name__ == "__main__":
    name = win32print.GetDefaultPrinter()
    print(f"Default Printer: {name}")
    
    # Try setting to A5 (11)
    print("--- Setting to A5 ---")
    set_paper_size(name, 11)
    
    # Check again
    hPrinter = win32print.OpenPrinter(name)
    dm = win32print.GetPrinter(hPrinter, 2)["pDevMode"]
    print(f"Verified Paper Size: {dm.PaperSize}")
    win32print.ClosePrinter(hPrinter)
    
    # Revert to A4 (9)
    print("--- Reverting to A4 ---")
    set_paper_size(name, 9)
