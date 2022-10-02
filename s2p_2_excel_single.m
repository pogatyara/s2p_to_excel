% ������ ������ ���������� � �����
% � ������� ���� .s2p

% ������ ���� ������ s2p
files = dir(fullfile('*.s2p') );
files = string({files.name}');

% �������� ������� �������� �������
% ����� .s2p
opts = detectImportOptions(files(1),'FileType','text');
% ����������� ��������� ������ � �����:
% ����� ������ � ���������
% % opts.Delimiter = {newline,'\t'};
% �������������� ���������� ��������
% � ���������, ������� �����
% (��������� � ���������)
opts.VariableNames = {...
    'freqs', ...
    'S11_db', 'S11_Ang',...
    'S21_db', 'S21_Ang',...
    'S12_db', 'S12_Ang',...
    'S22_db', 'S22_Ang'...
    };
% �������� ����� � ���������� �������
data = cell(numel(files),1);
% ���������� � ���������� .xlsx ������
for i=1:numel(files)
    % ���������� ���������� ������ 4� �����
    strings = split(fileread(files(i)),newline);
    % ������ � ������ ������� ��������
    % ����� .s2p � ����������� opts
    data{i} = readtable(files(i), opts);        
    % ������ ���������� ����������
    writecell(strings, ...
              strrep(files(i),'.s2p','.xlsx'),... ������.s2p �� .xlsx
                                              ... ��� �������� ����� ������
                                              ... � ��� �� ������
              'Sheet', 1, 'Range', 'A1');
    % ������ �������� ����������
    writetable(data{i}, ...
               strrep(files(i),'.s2p','.xlsx'), ...
               'Sheet', 1, 'Range', 'A5'); ...
end

