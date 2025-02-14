📌 Registro Automático de Última Alteração em Planilhas Excel

Este código VBA adiciona automaticamente a data e hora da última modificação em cada linha de uma planilha do Excel. Ele é útil para rastrear alterações em planilhas de controle de processos, garantindo que cada modificação seja registrada sem necessidade de intervenção manual.
🚀 Como Funciona?

Sempre que uma célula for alterada em uma linha, o código:
✔ Identifica a linha modificada.
✔ Registra a data e hora da alteração em uma coluna específica (por padrão, coluna J).
✔ Evita loops desnecessários ao ignorar alterações na própria coluna de registro.
🔧 Configuração e Uso

    Abrir o Editor VBA: Pressione ALT + F11 no Excel.
    Inserir o Código: Copie e cole o código no objeto da planilha desejada (Planilha1, Planilha2, etc.).
    Salvar como Macro-Enabled Workbook (.xlsm).
    Habilitar macros no Excel se necessário.

⚠ Observações

    OneDrive: Se o arquivo estiver no OneDrive, pode ser necessário abrir no Excel para Desktop e desativar o salvamento automático.
    Segurança: Caso o Excel bloqueie macros, vá até Arquivo > Opções > Central de Confiança e habilite macros.
    Personalização: Para registrar em outra coluna, basta alterar o valor de AlteracaoColuna.
