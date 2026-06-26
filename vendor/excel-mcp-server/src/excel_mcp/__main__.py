import sys

import typer

from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(help="Excel MCP Server")


def _stderr(message: str) -> None:
    print(message, file=sys.stderr)

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    try:
        run_sse()
    except KeyboardInterrupt:
        _stderr("\nShutting down server...")
    except Exception as e:
        _stderr(f"\nError: {e}")
        import traceback
        traceback.print_exc(file=sys.stderr)
    finally:
        _stderr("Service stopped.")

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    try:
        run_streamable_http()
    except KeyboardInterrupt:
        _stderr("\nShutting down server...")
    except Exception as e:
        _stderr(f"\nError: {e}")
        import traceback
        traceback.print_exc(file=sys.stderr)
    finally:
        _stderr("Service stopped.")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    try:
        run_stdio()
    except KeyboardInterrupt:
        _stderr("\nShutting down server...")
    except Exception as e:
        _stderr(f"\nError: {e}")
        import traceback
        traceback.print_exc(file=sys.stderr)
    finally:
        _stderr("Service stopped.")

if __name__ == "__main__":
    app() 
