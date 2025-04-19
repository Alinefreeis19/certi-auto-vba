# CertiAuto - Automação de Envio de Certificados em PDF via VBA

Este projeto tem como objetivo automatizar o envio de certificados em PDF para os supervisores de loja, contendo os certificados de todos os atendentes sob sua responsabilidade. A automação foi desenvolvida utilizando **VBA (Visual Basic for Applications)** e é integrada ao **Microsoft Excel** e **Outlook**, eliminando a necessidade de preenchimento e envio manual.

## Funcionalidades

- Geração automática de certificados em PDF
- Preenchimento dinâmico dos dados de cada atendente
- Envio automatizado dos arquivos via e-mail para o supervisor responsável
- Organização e padronização do processo de envio
- Redução de erros manuais e ganho de tempo na rotina operacional

## Como funciona

1. O código percorre uma base de dados com os nomes dos atendentes e seus respectivos supervisores.
2. Para cada atendente, gera um certificado em PDF com os dados preenchidos automaticamente.
3. Agrupa os certificados por supervisor e envia em anexo para o e-mail corporativo correspondente.
4. Finaliza o processo e exibe mensagem de sucesso.

## Tecnologias utilizadas

- VBA (Visual Basic for Applications)
- Microsoft Excel
- Microsoft Outlook
- Formato PDF

## Como usar

1. Abra o Excel e pressione `Alt + F11` para abrir o Editor VBA.
2. Importe o módulo `.bas` disponibilizado neste repositório.
3. Atualize a base de dados e os caminhos dos arquivos conforme necessário.
4. Execute a macro principal para iniciar o processo.

## Arquivos

- `envio_certificados.bas`: Código-fonte da automação em VBA.
- (Opcional) `modelo_certificado.xlsx`: Arquivo de exemplo com o modelo do certificado.

## Autor

Desenvolvido por [Aline Reis]  
[LinkedIn]([(https://www.linkedin.com/in/-aline-reis-0861b8157/]) | [GitHub](https://github.com/Alinefreeis19)

---
