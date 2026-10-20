import urllib.request
import ssl

def baixar_cids():
    urls = [
        "https://raw.githubusercontent.com/suportetitanium/cid10/master/cid10.json",
        "https://raw.githubusercontent.com/henriqueweiand/cid10/master/cid10.json"
    ]
    
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    print("A iniciar o download da base de dados do CID-10...")
    for url in urls:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=15) as response:
                dados = response.read().decode('utf-8')
                with open("cid10_local.json", "w", encoding="utf-8") as f:
                    f.write(dados)
            print("✔ Sucesso! O ficheiro 'cid10_local.json' foi criado na sua pasta.")
            print("Agora o seu sistema já não precisa de internet para buscar os CIDs!")
            return
        except Exception as e:
            print(f"Tentativa falhou no link: {url}. Erro: {e}")
            
    print("❌ Falha total. A sua rede está a bloquear completamente a ligação externa.")

if __name__ == "__main__":
    baixar_cids()
    input("\nPressione ENTER para sair...")