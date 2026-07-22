# -*- coding: utf-8 -*-
"""
Script de build para gerar EXE do Consulta Certidões ONR
Requer: pip install pyinstaller openpyxl
"""

import PyInstaller.__main__
import os
import shutil

# Usa Desktop como pasta de saída (mais confiável para antivírus)
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
output_folder = os.path.join(desktop, "BrejeiroCertidoes_Build")

# Limpa builds anteriores
for folder in ['build', 'dist', 'output']:
    try:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    except Exception:
        pass
try:
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
except Exception:
    pass

# Configurações do PyInstaller
PyInstaller.__main__.run([
    'consulta_certidoes_onr.py',
    '--name=BrejeiroCertidoes',
    '--onefile',
    '--windowed',
    '--distpath=' + output_folder,
    '--workpath=build',
    '--noupx',  # Desabilita UPX para evitar bloqueio
    '--noconfirm',
    '--icon=icon.ico' if os.path.exists('icon.ico') else '--icon=NONE',
    '--hidden-import=openpyxl',
    '--hidden-import=openpyxl.cell._writer',
    '--hidden-import=openpyxl.cell',
    '--hidden-import=openpyxl.workbook',
    '--clean',
])

# Cria arquivo version.txt na pasta output
version = "2.1.0"
with open(os.path.join(output_folder, 'version.txt'), 'w', encoding='utf-8') as f:
    f.write(version)

print(f"\n✅ Build concluído!")
print(f"📦 EXE gerado: {output_folder}/BrejeiroCertidoes.exe")
print(f"📄 Versão: {version}")
print(f"\n📋 Para distribuir:")
print(f"   1. Copie BrejeiroCertidoes.exe para a pasta de atualização do OneDrive")
print(f"   2. Copie version.txt para a mesma pasta")
print(f"   3. As máquinas irão baixar automaticamente ao abrir")
