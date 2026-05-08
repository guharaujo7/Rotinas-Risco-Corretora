# Rotinas-Risco-Corretora
Rotinas em python que auxilia algumas demandas da área de risco em corretoras.
  - SolutiongrupoBMF: Para utilizar troque url pelo link inicial do solutions, login e senha, tem que definir o grupo bmf que vai alterar, faça isso na linha 44 em value, veja o valor do grupo que quer alterar e altere no código! Ele gera log das alterações tem que criar um arquivo no diretorio chamado bmflimites_log.txt e pra rodar o codigo dos clientes, crie um arquivo chamado clientes.txt, um cliente por linha.
  - Monitoramento de limites: Realiza o calculo de rmkt em comitentes desenquadrados, utilizando spvd e etc. Utilize o excel e as informações enviadas pela B3.

cd "c:\Users\ghara\Desktop\projetos\itau"
.\build_mesa.ps1

    teste: roda em powershell pasta mesa
python -m pip install -r requirements-build.txt -q
python -m playwright install chromium
python -m PyInstaller mesa_itau.spec --noconfirm --clean