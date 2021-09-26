import logging

logger = logging.getLogger()
formatter = logging.Formatter("%(asctime)s %(message)s")
streamHandler = logging.StreamHandler()
streamHandler.setFormatter(formatter)
logger.addHandler(streamHandler)
logger.setLevel(logging.INFO)


def log(msg):
    logger.info(msg)
