import logging
import time
from selenium.common.exceptions import TimeoutException

def configurar_logging():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s",
                        handlers=[logging.StreamHandler()])

def pause(segundos):
    time.sleep(segundos)

def esperar_elemento(wait, by, selector, timeout=30):
    try:
        return wait.until(lambda d: d.find_element(by, selector))
    except TimeoutException:
        return None
