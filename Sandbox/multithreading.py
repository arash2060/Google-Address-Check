import threading
from queue import Queue
import requests
import bs4
import time

print_lock = threading.Lock()

def get_url(current_url):

    # with print_lock:
    #     print("\nStarting thread {}".format(threading.current_thread().name))
    res = requests.get(current_url)
    res.raise_for_status()

    current_page = bs4.BeautifulSoup(res.text,"html.parser")
    current_title = current_page.select('title')[0].getText()

    # with print_lock:
    #     print("{}\n".format(threading.current_thread().name))
    #     print("{}\n".format(current_url))
    #     print("{}\n".format(current_title))
    #     print("\nFinished fetching : {}".format(current_url))

def process_queue():
    while True:
        current_url = url_queue.get()
        get_url(current_url)
        url_queue.task_done()

url_queue = Queue()

url_list = ["https://www.google.com"]*5

for i in range(5):
    t = threading.Thread(target=process_queue)
    t.daemon = True
    t.start()

start = time.time()

for current_url in url_list:
    url_queue.put(current_url)

url_queue.join()

print(threading.enumerate())

print("Execution time = {0:.5f}".format(time.time() - start))
