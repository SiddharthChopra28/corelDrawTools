from threading import Thread
import pyautogui, time


class MainClass:
    def __init__(self):  
        i = 0
        Thread(target=lambda: A(), daemon=True).start()
        while 1:
            i += 1
            print(i)
            time.sleep(1)
      
# def func():
#     while 1:
#         print("a")
#         time.sleep(1)

class A:
    def __init__(self):
        while 1:
            self.x()
            time.sleep(1)
            
    def x(self):
        print("a")
  
if __name__ == '__main__':
    MainClass()