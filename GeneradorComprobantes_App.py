import _thread
#import sys
import os
import signal
from functions_GCA.window import startWindow
from pynput import mouse

if __name__ == '__main__':

    #Fail Safe
    def on_move(x, y):
        if(x==0 and y == 0):
            
            print("Corto")
            os.kill(os.getpid(), signal.SIGINT)
            #_thread.interrupt_main()
            #sys.exit()
    
    with mouse.Listener(on_move = on_move) as listener:
        startWindow()    #main
        _thread.interrupt_main()
        listener.join()
