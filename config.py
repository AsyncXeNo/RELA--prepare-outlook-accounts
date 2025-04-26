import sys
from datetime import datetime

from pyairtable import api
from dotenv import load_dotenv
from loguru import logger

load_dotenv()

timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') 
log_filename = f'logs/{timestamp}.log'

logger.remove()
logger.add(sys.stderr, format='<green>[{elapsed}]</green> <level>{level} > {message}</level>')
logger.add(log_filename, mode='w', format='[{elapsed}] {level} > {message}')

logger.debug('Logged initialized.')