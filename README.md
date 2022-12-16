# labelgen
Software independente para a geração de etiquetas com código QR, para a gestão de estoque e expedição.

# Instruções de instalação
1. Manter os scripts "main_file.py" e "pdfmerging.py" no mesmo diretório. Por exemplo: "C:\LabelGen\".
2. Dentro do diretório onde estão salvos os scripts do Passo 1, criar a estrutura de pastas: "<scripts>\ETIQUETAS\QRCODE" e "<scripts>\ETIQUETAS\PRODUCAO".
3. Compilar o "main_file.py" pelo método de sua preferência. Sugestão: criar um executável utilizando o pyinstaller (não se esqueça de fazer a mesma estrutura de pastas do Passo 2 no local onde será criada a build executável).
    3.1. Certifique-se de que você tem o módulo do PyInstaller instalado (Prompt de Comando > pip install pyinstaller)
    3.2. Vá até o diretório da pasta onde está o script "main_file.py" através do Prompt de Comando. Por exemplo: "cd C:\Users\Igor\Documents\Programas\LabelGen"
    3.3. Insira o comando: "pyinstaller main_file.py -F -w -n LabelGen_6_1"
    3.4. Após concluído, crie um .zip de tudo e distribua. O executável está dentro da pasta ".\dist"
