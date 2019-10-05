% % % % % % % % % % % % % % % % QUESTÃO 1 % % % % % % % % % % % % % % % %

% Questão 1 da tarefa referente à disciplina CIRCUITOS ELÉTRICOS II feita
% pelo grupo Composto por:

% Ana Paula Chaves Cabral - 20170024791
% Diana Bezerra Correia Lima - 20170017358
% Ruan Carlos Virginio dos Santos - 20170018711

% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 

clear all
clc
 
%% Nesta primeira seção, abrimos a planilha de entrada de dados do excel (Apenas formatos .xlsx e .xls)

[file, path] = uigetfile({'*.xlsx;*.xls'});

if(file==0) 
    disp 'Nenhum arquivo foi selecionado'
    return
end

dados = xlsread(strcat(path,file)); 

%% Nesta seção, vamos coletar os dados das colunas e transformar o número complexo separado em parte real e imaginária em um só

jj = sqrt(-1); %Complexo 

% Lendo os dados de cada coluna do arquivo e salvando nas variáveis

ZReal = dados(:,3); 
ZImag = dados(:,4);
VReal = dados(:,5);
VImag = dados(:,6);
IReal = dados(:,7);
IImag = dados(:,8);

% Transformando em um vetor complexo só 

Z = ZReal + jj*ZImag; 
Vs = VReal + jj*VImag;
Js = IReal + jj*IImag;

%% Nesta seção, coletamos o tamanho da matriz digitada pelo usuário e o número de ramos e nós do circuito

Tamanho = size(dados); % Tamanho da matriz de dados (linhas, colunas)
b = Tamanho(1); % O numero de linhas da matriz TAMANHO é o mesmo número de RAMOS do circuito

% Como cada ramo sai de um certo nó N e chega em outro nó N, essa seção
% verifica os valores escritos nas duas primeiras colunas da planilha e
% salva o maior valor encontrado na variável Nt

Nt = 0; % Número de NÓS

for i=1:b % Salva o maior número de nós analisando a coluna 1 da tabela
    ntemp = dados(i,1); % Coletando dados da coluna 1 (Nó de saída)
    
    if(ntemp > Nt)
        Nt = ntemp;
    end
    
    ntemp = dados(i,2); % Coletando dados da coluna 2 (Nó de entrada)

    if(ntemp > Nt)
        Nt = ntemp;
    end
end

%% Criando a Matriz de Incidência Ramo-Nó e sua Reduzida

% Criamos uma matriz nula do tamanho Número de Nós x Número de Ramos para
% receber a matriz incidência com "1" se é um nó de saída e "-1" se for um
% nó de entrada

Aa = zeros(Nt,b); 

for i=1:b % Alteração dos valores da matriz de acordo com a saida e entrada no nó
    Aa((dados(i,1)),i) =  1;
    Aa((dados(i,2)),i) = -1;
end

% Assim, teremos a matriz Incidência Ramo-Nó Aa:

Aa

% Agora, reduzimos essa matriz removendo sua última linha, que será
% considerada a do nó de referência do circuito

for i=1:(Nt -1)
    for j=1:b
        A(i,j) = Aa(i,j);
    end
end

% O que nos deixa, portanto, com a matriz de incidência Ramo-Nó reduzida A

A

%% Criação da Matriz Admitância de Ramo

% Vamos criar uma matriz quadrada nula Yb de tamanho Número de Ramos x Número de
% Ramos que irá receber em sua diagonal principal as admitâncias de cada
% ramo

Yb = zeros(b,b); 
for i=1:b
    Yb(i,i) = 1/(Z(i)); 
    % i = número do ramo, então o elemento (i,i) da matriz recebe a
    % admitância daquele ramo i
end

%% Resolução das equações da rede

Ye = A*Yb*A' % Matriz Admitância de Nó
Is = A*Yb*Vs - A*Js % Vetor Fontes de Corrente e Fontes de Tensão/Impedância do ramo que incidem no nó
E = inv(Ye)*Is % Vetor Tensões de Nó
V = (A')*E % Vetor Tensões de Ramo
J = (Js + (Yb*V) - (Yb*Vs)) % Vetor Correntes de Ramo

%% Enviando os dados das matrizes para um arquivo xls de saída

% Prepararemos a planilha de saída para receber os dados. Então, limparemos
% todos os dados anteriormentes salvos, exceto o cabeçalho que nos diz o
% que cada coluna nos diz.

% Name of the excel file
filename = 'C:\Users\santo_509teb3\Google Drive (ruan.santos@cear.ufpb.br)\Engenharia Elétrica\5º Período\Circuitos II\Projeto Circuitos II\Questão 1\Circuito.xls';
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


% A função num2str serve para transformar os valores complexos da matriz em
% uma string, pois sem essa função, só aparecia os valores reais de cada
% elemento dos vetores. Sem a função cellstr, cada caractere da string caía
% em uma coluna no excel. Com a função cellstr, a string fica em um
% elemento só na planilha

Vstr = num2str(V)
Vcell = cellstr(Vstr)

Jstr = num2str(J)
Jcell = cellstr(Jstr)

Estr = num2str(E)
Ecell = cellstr(Estr)

% A função xlswrite envia para a planilha CIRCUITO, as matrizes (V/J/E), 
% na folha RESULTADOS, para o elemento (A2/B2/C2) em diante.

xlswrite('Circuito', Vcell, 'Resultados', 'A2');
xlswrite('Circuito', Jcell, 'Resultados','B2');
xlswrite('Circuito', Ecell, 'Resultados', 'C2');

%% Algumas Referências 
% https://www.mathworks.com/matlabcentral/answers/7717-xlswrite-clear-cells-and-write-multiple-values
% https://www.mathworks.com/help/matlab/ref/uigetfile.html
% https://www.mathworks.com/help/matlab/ref/cellstr.html
