% % % % % % % % % % % % % % % % QUEST�O 1 % % % % % % % % % % % % % % % %

% Quest�o 1 da tarefa referente � disciplina CIRCUITOS EL�TRICOS II feita
% pelo grupo Composto por:

% Ana Paula Chaves Cabral - 20170024791
% Diana Bezerra Correia Lima - 20170017358
% Ruan Carlos Virginio dos Santos - 20170018711

% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 

clear all
clc
 
%% Nesta primeira se��o, abrimos a planilha de entrada de dados do excel (Apenas formatos .xlsx e .xls)

[file, path] = uigetfile({'*.xlsx;*.xls'});

if(file==0) 
    disp 'Nenhum arquivo foi selecionado'
    return
end

dados = xlsread(strcat(path,file)); 

%% Nesta se��o, vamos coletar os dados das colunas e transformar o n�mero complexo separado em parte real e imagin�ria em um s�

jj = sqrt(-1); %Complexo 

% Lendo os dados de cada coluna do arquivo e salvando nas vari�veis

ZReal = dados(:,3); 
ZImag = dados(:,4);
VReal = dados(:,5);
VImag = dados(:,6);
IReal = dados(:,7);
IImag = dados(:,8);

% Transformando em um vetor complexo s� 

Z = ZReal + jj*ZImag; 
Vs = VReal + jj*VImag;
Js = IReal + jj*IImag;

%% Nesta se��o, coletamos o tamanho da matriz digitada pelo usu�rio e o n�mero de ramos e n�s do circuito

Tamanho = size(dados); % Tamanho da matriz de dados (linhas, colunas)
b = Tamanho(1); % O numero de linhas da matriz TAMANHO � o mesmo n�mero de RAMOS do circuito

% Como cada ramo sai de um certo n� N e chega em outro n� N, essa se��o
% verifica os valores escritos nas duas primeiras colunas da planilha e
% salva o maior valor encontrado na vari�vel Nt

Nt = 0; % N�mero de N�S

for i=1:b % Salva o maior n�mero de n�s analisando a coluna 1 da tabela
    ntemp = dados(i,1); % Coletando dados da coluna 1 (N� de sa�da)
    
    if(ntemp > Nt)
        Nt = ntemp;
    end
    
    ntemp = dados(i,2); % Coletando dados da coluna 2 (N� de entrada)

    if(ntemp > Nt)
        Nt = ntemp;
    end
end

%% Criando a Matriz de Incid�ncia Ramo-N� e sua Reduzida

% Criamos uma matriz nula do tamanho N�mero de N�s x N�mero de Ramos para
% receber a matriz incid�ncia com "1" se � um n� de sa�da e "-1" se for um
% n� de entrada

Aa = zeros(Nt,b); 

for i=1:b % Altera��o dos valores da matriz de acordo com a saida e entrada no n�
    Aa((dados(i,1)),i) =  1;
    Aa((dados(i,2)),i) = -1;
end

% Assim, teremos a matriz Incid�ncia Ramo-N� Aa:

Aa

% Agora, reduzimos essa matriz removendo sua �ltima linha, que ser�
% considerada a do n� de refer�ncia do circuito

for i=1:(Nt -1)
    for j=1:b
        A(i,j) = Aa(i,j);
    end
end

% O que nos deixa, portanto, com a matriz de incid�ncia Ramo-N� reduzida A

A

%% Cria��o da Matriz Admit�ncia de Ramo

% Vamos criar uma matriz quadrada nula Yb de tamanho N�mero de Ramos x N�mero de
% Ramos que ir� receber em sua diagonal principal as admit�ncias de cada
% ramo

Yb = zeros(b,b); 
for i=1:b
    Yb(i,i) = 1/(Z(i)); 
    % i = n�mero do ramo, ent�o o elemento (i,i) da matriz recebe a
    % admit�ncia daquele ramo i
end

%% Resolu��o das equa��es da rede

Ye = A*Yb*A' % Matriz Admit�ncia de N�
Is = A*Yb*Vs - A*Js % Vetor Fontes de Corrente e Fontes de Tens�o/Imped�ncia do ramo que incidem no n�
E = inv(Ye)*Is % Vetor Tens�es de N�
V = (A')*E % Vetor Tens�es de Ramo
J = (Js + (Yb*V) - (Yb*Vs)) % Vetor Correntes de Ramo

%% Enviando os dados das matrizes para um arquivo xls de sa�da

% Prepararemos a planilha de sa�da para receber os dados. Ent�o, limparemos
% todos os dados anteriormentes salvos, exceto o cabe�alho que nos diz o
% que cada coluna nos diz.

% Name of the excel file
filename = 'C:\Users\santo_509teb3\Google Drive (ruan.santos@cear.ufpb.br)\Engenharia El�trica\5� Per�odo\Circuitos II\Projeto Circuitos II\Quest�o 1\Circuito.xls';
% Open Excel as a COM Automation server
Excel = actxserver('Excel.Application');
% Open Excel workbook
Workbook = Excel.Workbooks.Open(filename);
% Clear the content of the sheet
Workbook.Worksheets.Item('Resultados').Range('A2:C20').ClearContents  
% Now save/close/quit/delete
Workbook.Save;
Excel.Workbook.Close;
invoke(Excel, 'Quit');
delete(Excel)


% A fun��o num2str serve para transformar os valores complexos da matriz em
% uma string, pois sem essa fun��o, s� aparecia os valores reais de cada
% elemento dos vetores. Sem a fun��o cellstr, cada caractere da string ca�a
% em uma coluna no excel. Com a fun��o cellstr, a string fica em um
% elemento s� na planilha

Vstr = num2str(V)
Vcell = cellstr(Vstr)

Jstr = num2str(J)
Jcell = cellstr(Jstr)

Estr = num2str(E)
Ecell = cellstr(Estr)

% A fun��o xlswrite envia para a planilha CIRCUITO, as matrizes (V/J/E), 
% na folha RESULTADOS, para o elemento (A2/B2/C2) em diante.

xlswrite('Circuito', Vcell, 'Resultados', 'A2');
xlswrite('Circuito', Jcell, 'Resultados','B2');
xlswrite('Circuito', Ecell, 'Resultados', 'C2');

%% Algumas Refer�ncias 
% https://www.mathworks.com/matlabcentral/answers/7717-xlswrite-clear-cells-and-write-multiple-values
% https://www.mathworks.com/help/matlab/ref/uigetfile.html
% https://www.mathworks.com/help/matlab/ref/cellstr.html
