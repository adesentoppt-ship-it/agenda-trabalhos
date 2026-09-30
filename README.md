# Agenda de Trabalhos

App de telemóvel para agendar trabalhos, registar clientes (NIF, nome completo, morada) e os valores de cada serviço. O técnico usa a mesma app com o seu próprio PIN — não precisa de conta.

- **Onde ficam os dados:** numa folha Google Sheets na sua conta Google (não neste repositório).
- **Notificações:** cada trabalho cria um evento no seu Google Calendar com alarmes 60 e 15 min antes → o telemóvel avisa (app Google Calendar instalada e com a mesma conta).

## Instalar (uma vez, ~10 min)

1. Em sheets.google.com crie uma folha nova, chamada "Agenda de Trabalhos".
2. Menu **Extensões → Apps Script**.
3. Apague o conteúdo de `Código.gs` e cole o ficheiro `Codigo.gs` deste repositório.
   **Mude os PINs** no topo (`PIN_ADMIN` para si, `PIN_TECNICO` para o técnico).
4. Clique **+ → HTML**, chame-lhe `Index` e cole o conteúdo de `Index.html`.
5. Engrenagem (Definições do projeto) → Fuso horário **Europe/Lisbon**.
6. Escolha a função `configurar` e clique **Executar** → autorize o acesso à folha e ao calendário.
7. **Implementar → Nova implementação → Tipo: Aplicação Web**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
   → copie o endereço que termina em `/exec`.
8. Cole esse endereço em `config.js` neste repositório (ou na 1.ª entrada da app).

## Usar no telemóvel

Abra o endereço do GitHub Pages (ou o `/exec`) no telemóvel → menu do navegador → **Adicionar ao ecrã principal**. Envie o mesmo link ao técnico com o PIN dele.

- O técnico: vê a agenda, liga ao cliente, abre a morada no mapa, carrega em **Concluir** e coloca o valor.
- O dono (PIN admin): também pode apagar trabalhos. Aba **Valores**: total do mês, por técnico, por forma de pagamento e o que falta receber.

Se alterar o código no Apps Script: **Implementar → Gerir implementações → editar → Nova versão**.
