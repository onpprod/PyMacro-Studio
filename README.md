# PyMacro Studio

Aplicativo simples para Windows com interface gráfica para:

- Gravar macro de teclado e mouse
- Salvar/carregar macros em arquivo JSON
- Mapear uma tecla (ex.: `Numpad 7`) para executar uma macro salva

## Requisitos

- Python 3.10+
- Windows

## Instalação

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Execução

```powershell
python app.py
```

## Uso rápido

1. Digite um nome para a macro.
2. Clique em **Iniciar Gravação** e execute ações de teclado/mouse.
3. Clique em **Parar Gravação**.
4. Selecione a macro e clique em **Mapear Tecla**.
5. Pressione a tecla que deve disparar a macro (ex.: `Numpad 7`).
6. Para repetir continuamente, marque **Executar em loop** e ajuste **Loop (ms)** (mínimo `100ms`).
7. Use **Tecla de Parada do Loop** para definir a tecla que encerra o loop (padrão: `F8`).
8. Na tabela **Eventos da Macro**, edite o tempo (ms) de cada evento e aplique no evento selecionado ou em todos.

As macros e mapeamentos são salvos no arquivo `macros_db.json`.

## Observações

- Por padrão, o app grava apenas teclado. Marque **Gravar eventos do mouse** para incluir mouse.
- O painel de eventos mostra atraso individual e acumulado, permitindo ajuste fino da velocidade da macro.
- Para capturar e reproduzir entrada global no Windows, em alguns cenários o app pode precisar ser executado com permissões elevadas.
- Se houver antivírus/controle corporativo, eventos globais de teclado/mouse podem ser bloqueados.
