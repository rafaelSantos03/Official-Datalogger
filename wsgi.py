import os

# Garantir que dependências e diretório estejam corretos
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from conversordatalogger import app

# Expor 'app' para servidores WSGI como Gunicorn/Waitress
application = app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)