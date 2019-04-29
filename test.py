# for running tests
from win10toast import ToastNotifier

toaster = ToastNotifier()
toaster.show_toast(title='Warning', msg='CPU usage above 80%', duration=5, threaded=True)
