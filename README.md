# Instalacao

## 1 instalar o driver: AccessDatabaseEngine_X64.exe
## 2 Passo a Passo: Criar uma Tarefa no Windows para Executar um Arquivo na Inicialização

1. **Abrir o Agendador de Tarefas:**
   - Pressione `Win + R` para abrir a caixa de diálogo Executar.
   - Digite `taskschd.msc` e pressione Enter. Isso abrirá o Agendador de Tarefas do Windows.

2. **Criar uma Nova Tarefa:**
   - Dentro da pasta desejada ou diretamente em "Biblioteca do Agendador de Tarefas", clique com o botão direito e selecione "Criar Tarefa Básica".

3. **Nomear a Tarefa:**
   - Dê um nome significativo à sua tarefa e adicione uma descrição, se desejar. Clique em "Avançar".

4. **Configurar o Gatilho:**
   - Escolha a opção "Ao Iniciar o Computador" e clique em "Avançar".

5. **Selecionar a Ação:**
   - Escolha "Iniciar um Programa" e clique em "Avançar".

6. **Configurar o Programa a Ser Executado:**
   - Clique em "Procurar" e selecione o arquivo executável `s3.exe`. Clique em "Avançar".

7. **Concluir a Configuração:**
   - Revise as configurações da tarefa e, se estiverem corretas, clique em "Concluir".

8. **Configurar Segurança (Opcional):**
   - Se o arquivo exigir privilégios administrativos, vá para a pasta de tarefas, clique com o botão direito na tarefa recém-criada e escolha "Propriedades". Em "Geral", marque a opção "Executar com privilégios mais altos".

## Enviando planilha para o software e convertendo o banco de dados na planilha

1. Execute o `planilha-db.exe`
1. Escolha umas das opções que deseja fazer `Carregar a PLANILHA para o Banco de Dados` (selecione o xlsx) ou `Salvar Dados do Banco para a PLANILHA` (selecione a pasta para salvar o arquivo xlsx)
