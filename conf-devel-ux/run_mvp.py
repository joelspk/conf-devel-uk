import os

from mvp_app.main import create_server


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    server = create_server(host="0.0.0.0", port=port)
    print(f"Servidor ativo em http://127.0.0.1:{port}")
    server.serve_forever()
