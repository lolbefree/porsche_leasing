from win10toast import ToastNotifier


def my_notifier():
    toaster = ToastNotifier()
    toaster.show_toast("Електронний лист створено", "Електронний лист створено з вкладенням",
                       threaded=True, icon_path=None, duration=5)
