import win32com.client


def connect_to_solidworks():
    try:
        # اتصال به SolidWorks در حال اجرا
        sw = win32com.client.GetActiveObject("SldWorks.Application")

        print("Connected to SolidWorks successfully!")
        print("SolidWorks version:", sw.RevisionNumber)

        return sw

    except Exception as e:
        print("Could not connect to SolidWorks.")
        print("Error:", e)
        return None


if __name__ == "__main__":
    connect_to_solidworks()