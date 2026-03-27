from waitress import serve
from app import app

if __name__ == "__main__":
    print("-------------------------------------------")
    print("CatalogoPlus - SERVIDOR DE PRODUÇÃO ATIVO")
    print("O Túnel deve apontar para a porta: 8080")
    print("-------------------------------------------")

    serve(app, host='0.0.0.0', port=8080, threads=32, connection_limit=1000)