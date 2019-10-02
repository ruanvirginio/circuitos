% % % % % % % % % % % % % % % % QUESTÃO 2 % % % % % % % % % % % % % % % %

% Questão 2 da tarefa referente à disciplina CIRCUITOS ELÉTRICOS II feita
% pelo grupo Composto por:

% Ana Paula Chaves Cabral - 20170024791
% Diana Bezerra Correia Lima - 20170017358
% Ruan Carlos Virginio dos Santos - 20170018711

% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 

clear all
clc


% Como vamos trabalhar com os dados em laplace, precisaremos usar os
% valores como symbolic, tanto no tempo quanto na frequência

syms s t at ac faset fasec;

%% Nesta primeira seção, abrimos a planilha de entrada de dados do excel (Apenas formatos .xlsx e .xls)

[file, path] = uigetfile({'*.xlsx;*.xls'});
if(file==0) 
    disp 'Nenhum arquivo foi selecionado'
    return
end

dados = xlsread(strcat(path,file)); 

% Lendo os dados de cada coluna do arquivo e salvando nas variáveis

Nos = dados(:,1:2);
R = dados(:,3); 
H = dados(:,4);
F = dados(:,5);
ModT = sym(dados(:,6));
At = dados(:,7);
Bt = dados(:,8);
FaseT = sym(dados(:,9));
ModC = sym(dados(:,10));
Ac = dados(:,11);
Bc = dados(:,12);
FaseC = sym(dados(:,13));
i0 = dados(:,14);
v0 = dados(:,15);
TipoTensao = sym(dados(:,16));
TipoCorrente = sym(dados(:,17));

%% Nessa seção, encontramos o número de nós e ramos do circuito

% Como cada ramo sai de um certo nó N e chega em outro nó N, essa seção
% verifica os valores escritos nas duas primeiras colunas da planilha e
% salva o maior valor encontrado na variável Nt

Tamanho = size(dados); % Tamanho da matriz de dados (linhas, colunas)
b = Tamanho(1); % O numero de linhas da matriz TAMANHO é o mesmo número de RAMOS do circuito

Nt = 0; % Número de NÓS

for i=1:b % Salva o maior número de nós analisando a coluna 1 da tabela
    ntemp = Nos(i,1);
    
    if(ntemp > Nt)
        Nt = ntemp;
    end
    
    ntemp = Nos(i,2); %%coluna 2

    if(ntemp > Nt)
        Nt = ntemp;
    end
end

%% Criando a Matriz de Incidência Ramo-Nó 

% Criamos uma matriz nula do tamanho Número de Nós x Número de Ramos para
% receber a matriz incidência com "1" se é um nó de saída e "-1" se for um
% nó de entrada

Aa = zeros(Nt,b); 

for i=1:b % Alteração dos valores da matriz de acordo com a saida e entrada no nó
    Aa((Nos(i,1)),i) =  1;
    Aa((Nos(i,2)),i) = -1;
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

%% Nesta seção, começamos as trasnformações em Laplace:

% Passaremos todas as variáveis com que estamos trabalhando para symbolic.
% Quando os valores numéricos no tempo e na frequência forem iguais,
% multiplicaremos pela transformada do impulso, que é 1.

impt = dirac(t);
imp = laplace(impt);

% Dentre as diversas opções de transformadas, escolhemos trabalhar com o
% degrau, cuja transformada é:

ut = 1 + 0*t;
u = laplace(ut);

Vs = sym(zeros(b,1));
Js = sym(zeros(b,1));

Cossenot = sym(zeros(b,1));
Cossenoc = sym(zeros(b,1));

Senoc = sym(zeros(b,1));
Senot = sym(zeros(b,1));


for n = 1:b
    if (TipoTensao(n) == 1)
        Vs(n) = ModT(n)*imp;
    end
    
    if (TipoTensao(n) == 2)
        ut = 1 + 0*t;
        u = laplace(ut);
        Vs(n) = ModT(n)*u;
    end
    
    if (TipoTensao(n) == 3)
        at = At(n);
        exptt = exp(-at*t);
        expst = laplace(exptt,s);
        Vs(n) = ModT(n)*expst;
    end
    
    if (TipoTensao(n) == 4)
        at = At(n);
        Senot(n) = laplace(sin(at*t + faset));
        Vs(n) = ModT(n)*Senot;
    end
    
    if (TipoTensao(n) == 5)
        at = At(n);
        faset = FaseT(n);
        Cossenot(n) = laplace(cos(at*t + faset));
        Vs(n) = ModT(n)*Cossenot(n);
    end
    
    if (TipoTensao(n) == 6)
        at = At(n);
        bt = Bt(n);
        faset = FaseT(n);
        exptt = exp(-at*t);
        expst = laplace(exptt,s);
        Senot(n) = laplace(sin(bt*t + faset));
        Vs(n) = ModT(n)*expst*Senot(n);
    end
    
    if (TipoTensao(n) == 7)
        at = At(n);
        bt = Bt(n);
        faset = FaseT(n);
        exptt = exp(-at*t);
        expst = laplace(exptt,s);
        Cossenot(n) = laplace(cos(bt*t + faset));
        Vs(n) = ModT(n)*expst*Cossenot(n);
    end
    
% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %

    if (TipoCorrente(n) == 1)
        Js(n) = ModC(n)*imp;
    end
    
    if (TipoCorrente(n) == 2)
        ut = 1 + 0*t;
        u = laplace(ut);
        Js(n) = ModC(n)*u;
    end
    
    if (TipoCorrente(n) == 3)
        ac = Ac(n);
        exptc = exp(-ac*t);
        expsc = laplace(exptt,s);
        Js(n) = ModC(n)*expsc;  
    end
    
    if (TipoCorrente(n) == 4)
        ac = Ac(n);
        fasec = FaseC(n);
        Senoc(n) = laplace(sin(ac*t + fasec));
        Js(n) = ModC(n)*Senoc(n);
    end
    
    if (TipoCorrente(n) == 5)
        ac = Ac(n);
        fasec = FaseC(n);
        Cossenoc(n) = laplace(cos(ac*t + fasec));
        Js(n) = ModC(n)*Cossenoc(n);
    end
    
    if (TipoCorrente(n) == 6)
        ac = Ac(n);
        bc = Bc(n);
        fasec = FaseC(n);
        exptc = exp(-ac*t);
        expsc = laplace(exptc,s);
        Senoc(n) = laplace(sin(bc*t + fasec));
        Js(n) = ModC(n)*expsc*Senoc(n);
    end
    
    if (TipoCorrente(n) == 7)
        ac = Ac(n);
        bc = Bc(n);
        fasec = FaseC(n);
        exptc = exp(-ac*t);
        expsc = laplace(exptc,s);
        Cossenoc(n) = laplace(cos(bc*t + fasec));
        Js(n) = ModC(n)*expsc*Cossenoc(n);
    end
    
end

A = A*imp;

% O for abaixo cria os vetores das impedâncias de cada um dos elementos que
% podem estar em cada ramo. Ainda, ela cria as fontes de tensão que são
% geradas na transformada a partir das condições iniciais no capacitor e no
% indutor.

for k = 1:b
    
    Zl(k,1) = H(k)*(1/u);
    Zc(k,1) = u/F(k);
    Zr(k,1) = R(k)*imp;
    Vcap(k,1) = v0(k)*u;
    Vind(k,1) = - i0(k)*H(k);
    
    if F(k) == 0
        Zc(k,1) = 0;
    end
    
end

Vs

%% Matrizes Fontes de tensão, fontes de corrente e impedâncias no ramo

% A matriz fontes de tensão independentes no ramo recebe as tensões geradas
% pelo capacitor e indutor. A matriz impedância recebe a soma das
% impedâncias geradas por todos os elementos também.

Vs = Vs + Vcap*imp + Vind*imp;

Z = Zr + Zc + Zl;

%% Criação da Matriz Admitância de Ramo

% Vamos criar uma matriz quadrada nula Yb de tamanho Número de Ramos x Número de
% Ramos que irá receber em sua diagonal principal as admitâncias de cada
% ramo

Yb = sym(zeros(b,b)); 
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

VRamo = ilaplace(V)
ENo = ilaplace(E)
JRamo = ilaplace(J)

%% Passando pra planilha

% Name of the excel file
filename = 'C:\Users\santo_509teb3\Google Drive (ruan.santos@cear.ufpb.br)\Engenharia Elétrica\5º Período\Circuitos II\Projeto Circuitos II\Questão 2\Circuito_Laplace.xls';
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

% Aqui temos mais e mais problemas. Não conseguimos criar o arquivo de
% saída diretamente com os dados em simbólico. Também não foi possível
% transformá-lo direto para um dado numérico através da função "double",
% então a saída encontrada foi transformar em um char, dividir o elemento
% único (que já estava previamente separada por vírgulas os elementos que 
% desejávamos quando em CHAR) em três strings separadas. Ainda, eliminamos
% o texto que não desejávamos exibir, e pegamos a transposta (pois a matriz
% original estava em linha) para exibir finalmente na planilha de saída.

Vchar = char(VRamo);
Vsplit = strsplit(Vchar,', ');
Vcell = cellstr(Vsplit);
Erase1 = regexprep(Vcell,'matrix([','','ignorecase');
Erase2 = regexprep(Erase1,'])','','ignorecase');
Vcell = Erase2';

Echar = char(ENo);
Esplit = strsplit(Echar,', ');
Ecell = cellstr(Esplit);
Erase3 = regexprep(Ecell,'matrix([','','ignorecase');
Erase4 = regexprep(Erase3,'])','','ignorecase');
Ecell = Erase4';

Jchar = char(JRamo);
Jsplit = strsplit(Jchar,', ');
Jcell = cellstr(Jsplit);
Erase5 = regexprep(Jcell,'matrix([','','ignorecase');
Erase6 = regexprep(Erase5,'])','','ignorecase');
Jcell = Erase6';

xlswrite('Circuito_Laplace', Vcell, 'Resultados', 'A2');
xlswrite('Circuito_Laplace', Jcell, 'Resultados', 'B2');
xlswrite('Circuito_Laplace', Ecell, 'Resultados', 'C2');

TensaoRamo = VRamo(5);
tempo = [0:0.0001:0.4];
corramo = double(subs(TensaoRamo,symvar(TensaoRamo),tempo));

plot(tempo,corramo);

%% Algumas Referências 
% https://www.mathworks.com/matlabcentral/answers/7717-xlswrite-clear-cells-and-write-multiple-values
% https://www.mathworks.com/help/matlab/ref/uigetfile.html
% https://www.mathworks.com/help/symbolic/laplace.html
% https://www.mathworks.com/help/symbolic/ilaplace.html
% https://uk.mathworks.com/matlabcentral/answers/356617-how-to-remove-a-some-part-of-a-string-from-a-string
% Além de vários e vários tópicos no fórum do mathworks