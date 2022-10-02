% скрипт должен находитьс€ в папке
% с файлами типа .s2p

% список имен файлов s2p
files = dir(fullfile('*.s2p') );
files = string({files.name}');

% создание объекта настроек импорта
% файла .s2p
opts = detectImportOptions(files(1),'FileType','text');
% ќбозначение делителей текста в файле:
% нова€ строка и табул€ци€
% % opts.Delimiter = {newline,'\t'};
% переименование заголовков столбцов
% в заголовки, которые нужны
% (нуждаетс€ в доработке)
opts.VariableNames = {...
    'freqs', ...
    'S11_db', 'S11_Ang',...
    'S21_db', 'S21_Ang',...
    'S12_db', 'S12_Ang',...
    'S22_db', 'S22_Ang'...
    };
% создание €чеек с табличными данными
data = cell(numel(files),1);
% заполнение и сохранение .xlsx файлов
for i=1:numel(files)
    % нецифрова€ информаци€ первых 4х строк
    strings = split(fileread(files(i)),newline);
    % запись в €чейку таблицы значений
    % файла .s2p с настройками opts
    data{i} = readtable(files(i), opts);        
    % запись нечисловой информации
    writecell(strings, ...
              strrep(files(i),'.s2p','.xlsx'),... замена.s2p на .xlsx
                                              ... дл€ создани€ файла эксель
                                              ... с тем же именем
              'Sheet', 1, 'Range', 'A1');
    % запись числовой информации
    writetable(data{i}, ...
               strrep(files(i),'.s2p','.xlsx'), ...
               'Sheet', 1, 'Range', 'A5'); ...
end

