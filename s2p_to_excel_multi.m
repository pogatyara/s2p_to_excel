% для работы нужен RF Toolbox
clear;

%--------параметры--------

% показывать фазы в градусах
showAnglesInDegrees = true;

% коэффициент пересчета из ГГц в Гц
freqCoef = 1E+9;   

% индексы рассмативаемых частот
freqMin = 1.2; % ГГц
freqMax = 1.4; % ГГц

% сохранить массивы в таблице
saveTable = true;
tableFormat = '.xlsx';
% порядок сохраения листов в таблице эксель
saveList = {...
    'S21dB',   'S21Ang',...
    'S11dB',   'S11Ang',...
    'S12dB',   'S12Ang',...
    'S22dB',   'S22Ang',...
    'VSWR',    'DN'...
    };

%-------------------------

% полный путь к текущей директории
cDir = cd;

% название выходной таблицы
outputTableName = split(cDir,'\');
outputTableName = char(outputTableName(end));

% список файлов .s2p
fileDirs = dir(fullfile('*.s2p'));

% сортировка имен файлов (углов) по возрастанию
fileNames = string({fileDirs.name});
angleArray = double(strrep(fileNames,'.s2p',''));
angleArray = sort(angleArray);
fileNames = string(angleArray) + ".s2p";

% ширина массивов
width = size(fileNames);
width = width(1,2);

% определение сетки частот
S = sparameters(fileNames(1,1));
nFreqMin = find(S.Frequencies == freqMin * freqCoef);
nFreqMax = find(S.Frequencies == freqMax * freqCoef);
freqRange = [nFreqMin : nFreqMax];

% запись сетки частот
freqs = S.Frequencies(freqRange) / freqCoef;

% длина массивов
length = size(freqs);

% имена извлекаемых переменных
SParameterVarNames = {...
    'S11',      'S12',      'S21',      'S22';...
    'S11dB',    'S12dB',    'S21dB',    'S22dB';...
    'S11Ang',   'S12Ang',   'S21Ang',   'S22Ang';...
    }; 

for i=1:width
    % S-параметры из .s2p
    S = sparameters(fileNames(i));

    for ik = 1:4
        % индексы Smn
        m = mod(ik,2) + floor(ik / 2);        
        n = 1 + (mod(ik,2) == 0);                 
        
        % имя переменной Smn
        Smn = SParameterVarNames{1, ik};
        % присвоение значений соответствующего S-параметра
        assignin('base', Smn, S.Parameters(m, n, freqRange))
        % изменение размерности
        assignin('base', Smn, reshape(evalin('base', Smn),[length, 1]));
        
        % имя переменной SmndB
        SmndB = SParameterVarNames{2, ik};  
        assignin('base', SmndB, 20 * log10(abs(evalin('base', Smn))));
        
        % имя переменной SmnAng
        SmnAng = SParameterVarNames{3, ik}; 
        if showAnglesInDegrees
            assignin('base', SmnAng, rad2deg(angle(evalin('base', Smn))));     
        else
            assignin('base', SmnAng, angle(evalin('base', Smn)));
        end
        
        if i==1
            % инициализация таблиц
            assignin('base', [SmndB 'Array'], evalin('base', SmndB));
            assignin('base', [SmnAng 'Array'], evalin('base', SmnAng));        
            
            % инициализация таблицы КСВ
            if m == 1 && n == 1 
                VSWRArray = (1 + 10.^(0.1*S11dB))./(1 - 10.^(0.1*S11dB));
            end
        else
            % добавление столбца значений 
            assignin('base', [SmndB 'Array'], ...
                cat(2, evalin('base', [SmndB 'Array']), evalin('base', SmndB)));       
            assignin('base', [SmnAng 'Array'], ...
                cat(2, evalin('base', [SmnAng 'Array']), evalin('base', SmnAng)));
            
            % добавление столбца значений КСВ 
            if m==1 && n==1
            VSWRArray = cat(2, VSWRArray, (1+10.^(0.1*S11dB))./(1-10.^(0.1*S11dB)));        
            end
        end
    % удаление лишних массивов (для удобства в рабочем пространстве)
     clear(SParameterVarNames{:});
    end
end

% расчет нормированной ДН
DNArray = zeros([length(1),width]);
for row = 1:length
  for col = 1:width
      DNArray(row,col) = S21dBArray(row,col) - max(S21dBArray(row,:));
  end
end

if saveTable
    % сохранение массивов
    for ik = 1:numel(saveList)

        % имя листа таблицы, согласно списку сохранения
        sheetName = saveList{ik};

        writematrix(...
            evalin('base',[sheetName 'Array']),...
            [outputTableName tableFormat],...
            'Sheet', sheetName,...
            'Range', 'B2');
        writematrix(...
            freqs,...
            [outputTableName tableFormat],...
            'Sheet', sheetName,...
            'Range', 'A2');        
        writematrix(...
            angleArray,...
            [outputTableName tableFormat],...
            'Sheet', sheetName,...
            'Range', 'B1');         
    end
end