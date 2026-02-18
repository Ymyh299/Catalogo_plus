from waitress import serve
from app import app  # Isso importa o objeto 'app' do seu arquivo app.py

if __name__ == "__main__":
    print("-------------------------------------------")
    print("CatalogoPlus - SERVIDOR DE PRODUÇÃO ATIVO")
    print("O Túnel deve apontar para a porta: 8080")
    print("-------------------------------------------")

    # O Waitress roda como HTTP internamente. 
    # O Playit.gg cuidará do HTTPS para as funcionárias no Sul.
    serve(app, host='0.0.0.0', port=8080, threads=4)