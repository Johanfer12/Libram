from playsound import playsound
from datetime import datetime
import time
start = time.time()

time.sleep(10)  # or do something more productive

done = time.time()
elapsed = done - start
print(elapsed)



playsound('cat.wav')
dateTimeObj = datetime.now()
print(dateTimeObj)