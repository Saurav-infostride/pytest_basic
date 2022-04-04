import logging

logging.basicConfig(filename=r"C://Users//SauravSharma//Pytest//classes//selenium_logs//tests.log",
                                    format='%(asctime)s: %(levelname)s: (levelname)s: %(message)s',
                                    level = logging.DEBUG
                                    )

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

logger.debug("This is a debug message")
logger.info("This is a info message")
logger.warning("This is a warning message")
logger.error("This is a error message")
logger.critical("This is a critical message")