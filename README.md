# Hub de Ferramentas & Aplicações Contábeis

Plataforma integrada de utilitários contábeis, gestão de documentos, automação de processos e auditoria fiscal/bancária.

---

## 🚀 Módulos Integrados

### 1. 📄 GhostPDF Studio (`/pdf-studio`)
- **Compressão Inteligente:** Motor Ghostscript 10 com presets otimizados (Screen, Ebook, Prepress, Default).
- **OCR Automatizado:** Reconhecimento óptico de caracteres com Tesseract OCR integrado para tornar PDFs pesquisáveis.
- **Fusão & Mesclagem:** Unificação rápida de múltiplos documentos PDF.
- **Divisão & Extração:** Separação por páginas ou faixas personalizadas.
- **Imagens para PDF:** Conversão e ordenação de fotos (JPG, PNG, TIFF) em documento PDF único.
- **Pasta Monitorada (Watch Folder):** Serviço em segundo plano que comprime e arquiva documentos automaticamente na rede ou disco local.

### 2. 🛡️ Central de Certificados Digitais (`/certificados`)
- **Sincronização Incremental:** Robô que varre as pastas de rede (`T:\Geral\Publica\CERTIFICADOS DIGITAIS`), descriptografa arquivos `.pfx` com base nas senhas da pasta e alimenta a planilha oficial de vencimentos.
- **Cache Inteligente:** Continua exatamente da última verificação, analisando apenas arquivos novos ou alterados.
- **Ordenação Dinâmica:** Colunas ordenáveis (Nº Empresa, Razão Social, CNPJ, Vencimento e Status/Dias Restantes).
- **Validador Local de PFX:** Arraste e solte certificados no navegador para inspeção instantânea.

### 3. 🧮 Conciliação Bancária (`/conciliacao`)
- **Importação de Extratos em PDF:** Leitura automatizada de extratos bancários de diversos bancos.
- **Lançamentos & Resumo:** Apuração detalhada mês a mês de entradas, saídas e saldos.
- **Controle de Aplicações & Rendimentos:** Acompanhamento de fundos, CDBs e rendimentos tributáveis/isentos.
- **Auditoria de Pendências:** Identificação de divergências contábeis e categorização de despesas.
- **Exportação:** Geração de relatórios em Excel (`.xlsx`) e arquivos formatados em `.txt` para integração com sistemas contábeis.

---

## 🛠️ Tecnologias Utilizadas
- **Backend:** Python 3.14 + Flask
- **Processamento de PDF:** Ghostscript 10, PyPDF
- **OCR:** Tesseract OCR 5.5 + PyTesseract
- **Criptografia & Certificados:** `cryptography` (PKCS#12)
- **Manipulação de Planilhas:** OpenPyXL
- **Frontend:** TailwindCSS, Lucide Icons, JavaScript Moderno

---

## 💻 Execução Local

```bash
# Instalar dependências
pip install flask pypdf pillow pytesseract openpyxl cryptography

# Iniciar servidor
python app.py
```

Acesse no navegador:
- Local: `http://127.0.0.1:5000`
- Rede Interna: `http://<IP-DO-SERVIDOR>:5000`
