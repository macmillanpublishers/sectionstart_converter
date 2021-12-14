import logging
import functools
import time

# initialize logger
logger = logging.getLogger(__name__)

# # top level function is just to allow for decorator parameters
def retry(max_retries=3, sleeptime=1):
    # decorator function taking function as argument
    def retry_core(func):
        # functools allows for use of relative constants (like __name__) in higher order functions
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            count = 0
            success = False
            while count < max_retries and success == False:
                count = count + 1
                # try to run the function, retries kickoff on exception
                try:
                    func(*args, **kwargs)
                    success=True
                except:
                    # run retry
                    logger.info('* Caught exception from "{}". Retrying, attempt {} of {} (waiting {} second(s) between attempts)'
                        .format(func.__name__, count, max_retries, sleeptime))
                    time.sleep(sleeptime)
                    # max_retries were unsuccessful, raise error
                    if count == max_retries:
                        logger.warn('** {} retries unsuccessful, raising exception for "{}".'.format(max_retries, func.__name__))
                        raise
        return wrapper
    return retry_core

def benchmark(func):
    """
    A decorator that prints the time a function takes
    to execute.
    """
    import time
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            print ("s")
            t = time.clock()
            p
            res = func(*args, **kwargs)
            print (func.__name__, " elapsed: {}".format(time.clock()-t))
            return res
        except:
            print ("q")
            raise
    return wrapper

def debug_logging(func):
    """
    A decorator that logs the activity of the script.
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        res = func(*args, **kwargs)
        logstring = "* running func: {}, args: {}, kwargs: {}".format(func.__name__, args, kwargs)
        # print logstring
        logging.debug(logstring)
        return res
    return wrapper

def counter(func):
    """
    A decorator that counts and prints the number of times a function has been executed
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        wrapper.count = wrapper.count + 1
        res = func(*args, **kwargs)
        print ('{0} has been used: {1}x'.format(func.__name__, wrapper.count))
        return res
    wrapper.count = 0
    return wrapper
