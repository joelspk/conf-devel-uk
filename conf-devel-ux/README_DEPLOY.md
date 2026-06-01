# Conferencia Bancaria Web

## O que mudou

- A aplicacao agora funciona como web app compartilhavel.
- O usuario pode:
  - enviar arquivos em lote;
  - enviar uma pasta inteira pelo navegador;
  - enviar um ZIP com o lote.

## Importante

Em ambiente web publico, o servidor nao consegue abrir uma pasta do computador do usuario por caminho local.
Por isso, a alternativa correta e:

- enviar a pasta pelo navegador;
- ou enviar um arquivo ZIP.

## Deploy rapido no Render

1. Suba estes arquivos em um repositorio Git.
2. No Render, crie um novo `Web Service`.
3. Aponte para esse repositorio.
4. O Render vai usar:
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`
5. Quando publicar, voce recebera uma URL publica para todos os usuarios.

## Rodar localmente

```bash
pip install -r requirements.txt
python run_mvp.py
```

## URL local

```text
http://127.0.0.1:8000
```
