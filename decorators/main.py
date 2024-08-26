

from anyio import current_time
from fastapi import FastAPI, Request, Response
import uvicorn

from logger import log_request

app = FastAPI()

@app.get("/hello_world")
@log_request
async def say_hello(request: Request):
    # Simulate processing the user info
    print("returning hello world"+ str(current_time))
    return Response(status_code=200, content="Hello, World! Time is: " +  str(current_time))

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)