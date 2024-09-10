# This module contains a decorator function that logs the request and response of a given function.
import time
import concurrent.futures
from functools import wraps
from fastapi import Request
from fastapi.responses import JSONResponse
# Initialize a ThreadPoolExecutor
executor = concurrent.futures.ThreadPoolExecutor(max_workers=10)

def log_request(func):
    """
    Decorator function that logs the request and response of a given function.
    Args:
      func: The function to be decorated.
    Returns:
      The decorated function.
    """
    
    @wraps(func)
    async def wrapper(*args, **kwargs):
        # Extract the request object from the arguments
        request = kwargs.get("request")
        # if request:
        #     # Submit the logging function to the executor
        #     executor.submit(log_to_server, request)
        
        # Call the original function and get the response
        response = await func(*args, **kwargs)
        
        # Log the response
        if request and response:
            executor.submit(log_response, request, response)
        
        return response

    return wrapper

def log_to_server(request: Request):
    """
    Logs the given request to the server.
    Args:
        request (Request): The request object to be logged.
    Returns:
        None
    """
    # Simulate a time-consuming logging operation
    time.sleep(5)
    print(f"Logged request: {request.url}")

def log_response(request: Request, response: JSONResponse):
    """ 
    Logs the given response to the server.
    Args:
        request (Request): The request object.
        response (JSONResponse): The response object.
    Returns:
        None
    """
    # Simulate a time-consuming logging operation
    time.sleep(5)
    print(f"Logged response for {request.url}: {response.status_code}")