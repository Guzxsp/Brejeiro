# Sistema de Atualização Automática - Certidões ONR

## 📦 Como Gerar o EXE

### Pré-requisitos
```bash
pip install pyinstaller openpyxl
```

### Gerar o Executável
```bash
python build_exe.py
```

Isso criará:
- `dist/CertidoesONR.exe` - Executável final
- `dist/version.txt` - Arquivo de versão

## 🔄 Sistema de Atualização Automática

### Configuração da Pasta de Atualização

O sistema usa uma pasta sincronizada via OneDrive para distribuir atualizações:

**Caminho configurado no código:**
```
C:\Users\ANPSOJA3\OneDrive - BREJEIRO\Atualizacoes_CertidoesONR
```

**Conteúdo da pasta:**
- `CertidoesONR.exe` - Versão mais recente do executável
- `version.txt` - Número da versão (ex: 2.1.0)

### Como Publicar uma Atualização

1. **Gere o novo EXE:**
   ```bash
   python build_exe.py
   ```

2. **Atualize o número da versão no código:**
   - Edite `consulta_certidoes_onr.py`
   - Altere `APP_VERSION = "2.1.0"` para a nova versão
   - Ex: `APP_VERSION = "2.2.0"`

3. **Copie os arquivos para a pasta de atualização:**
   - Copie `dist/CertidoesONR.exe` para a pasta do OneDrive
   - Copie `dist/version.txt` para a pasta do OneDrive
   - Aguarde o OneDrive sincronizar

4. **Pronto!** As máquinas irão detectar a atualização automaticamente ao abrir.

### Como Funciona

1. **Ao abrir o aplicativo:**
   - Verifica versão local (`version.txt` na pasta do EXE)
   - Verifica versão remota (pasta do OneDrive)
   - Se houver diferença, mostra diálogo de atualização

2. **Se usuário aceitar atualizar:**
   - Baixa o novo EXE da pasta do OneDrive
   - Faz backup do EXE atual
   - Substitui pelo novo
   - Reinicia o aplicativo

3. **Se usuário recusar:**
   - Abre normally com a versão antiga

### Para Outras Máquinas

**Instalação inicial:**
1. Copie `CertidoesONR.exe` para a máquina
2. Copie `version.txt` para a mesma pasta
3. Crie atalho na área de trabalho se desejar

**Atualizações subsequentes:**
- Automáticas! Basta abrir o aplicativo
- Ele irá detectar e oferecer atualização

## ⚙️ Configurações

No arquivo `consulta_certidoes_onr.py`, você pode alterar:

```python
APP_VERSION = "2.1.0"  # Versão atual do código
UPDATE_FOLDER = r"C:\Users\ANPSOJA3\OneDrive - BREJEIRO\Atualizacoes_CertidoesONR"  # Pasta de atualizações
```

## 🔧 Solução de Problemas

**Atualização não é detectada:**
- Verifique se `version.txt` existe na pasta do EXE
- Verifique se a pasta do OneDrive está sincronizada
- Verifique se o caminho `UPDATE_FOLDER` está correto

**Erro ao aplicar atualização:**
- Verifique permissões na pasta do EXE
- Verifique se o EXE não está em uso
- Feche o aplicativo antes de atualizar manualmente

**Build falha:**
- Instale dependências: `pip install pyinstaller openpyxl`
- Verifique se `icon.ico` existe (opcional)
