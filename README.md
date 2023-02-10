# Macro-Excluir-RDA-Automatico
VBA EXCEL com SAP GUI

Esta macro exclui RDAs com uma linha, da transação ME53N do SAP GUI usando o VBA do Excel.
Essa macro por mais simples que seja, ajuda bastante quando tem uma grande quantidade de RDAs para se excluir, pois a exclusão de RDA é algo bem repetitivo.

Passo a passo de como usar:
1. Abra o SAP.
2. Digite todas as RDAs que você deseja excluir abaixo do campo "RDAs para excluir"(Digite cada número da RDA em uma linha diferente para que ele possa excluir todas as RDAs desejadas).
3. Clique no botão "Excluir RDAs.".
4. Preencha o campo Login e Senha, com o seu login e senha do SAP ou clique em "Já estou logado", caso já esteja logado no SAP.
5. Clique em OK.

Obs:
A chance de a macro apresentar algum erro é pequena, mas é possível, pelo fato do SAP poder ser usado com várias configurações prévias diferentes.
Caso aconteça algum erro, eu escrevi o que cada parte do código está fazendo no VBA, então caso for necessário alguma mudança, é simples de se resolver.
