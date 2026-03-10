#!/usr/bin/env python3
"""Converte logo.png em logo.ico para o build."""

from PIL import Image

print("Convertendo logo.png para logo.ico...")

try:
    img = Image.open("logo.png")
    img = img.resize((256, 256), Image.Resampling.LANCZOS)
    img.save(
        "logo.ico",
        format="ICO",
        sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)],
    )
    print("Sucesso: arquivo 'logo.ico' criado.")
    print("Este arquivo sera usado como icone da barra de tarefas.")
except FileNotFoundError:
    print("Erro: 'logo.png' nao encontrado na pasta atual.")
    print("Certifique-se de estar na pasta correta.")
except Exception as e:
    print(f"Erro ao converter: {e}")
