% ��� ������ ����� RF Toolbox
clear;

%--------���������--------

% ���������� ���� � ��������
showAnglesInDegrees = true;

% ����������� ��������� �� ��� � ��
freqCoef = 1E+9;   

% ������� �������������� ������
freqMin = 1.2; % ���
freqMax = 1.4; % ���

% ��������� ������� � �������
saveTable = true;
tableFormat = '.xlsx';
% ������� ��������� ������ � ������� ������
saveList = {...
    'S21dB',   'S21Ang',...
    'S11dB',   'S11Ang',...
    'S12dB',   'S12Ang',...
    'S22dB',   'S22Ang',...
    'VSWR',    'DN'...
    };

%-------------------------

% ������ ���� � ������� ����������
cDir = cd;

% �������� �������� �������
outputTableName = split(cDir,'\');
outputTableName = char(outputTableName(end));

% ������ ������ .s2p
fileDirs = dir(fullfile('*.s2p'));

% ���������� ���� ������ (�����) �� �����������
fileNames = string({fileDirs.name});
angleArray = double(strrep(fileNames,'.s2p',''));
angleArray = sort(angleArray);
fileNames = string(angleArray) + ".s2p";

% ������ ��������
width = size(fileNames);
width = width(1,2);

% ����������� ����� ������
S = sparameters(fileNames(1,1));
nFreqMin = find(S.Frequencies == freqMin * freqCoef);
nFreqMax = find(S.Frequencies == freqMax * freqCoef);
freqRange = [nFreqMin : nFreqMax];

% ������ ����� ������
freqs = S.Frequencies(freqRange) / freqCoef;

% ����� ��������
length = size(freqs);

% ����� ����������� ����������
SParameterVarNames = {...
    'S11',      'S12',      'S21',      'S22';...
    'S11dB',    'S12dB',    'S21dB',    'S22dB';...
    'S11Ang',   'S12Ang',   'S21Ang',   'S22Ang';...
    }; 

for i=1:width
    % S-��������� �� .s2p
    S = sparameters(fileNames(i));

    for ik = 1:4
        % ������� Smn
        m = mod(ik,2) + floor(ik / 2);        
        n = 1 + (mod(ik,2) == 0);                 
        
        % ��� ���������� Smn
        Smn = SParameterVarNames{1, ik};
        % ���������� �������� ���������������� S-���������
        assignin('base', Smn, S.Parameters(m, n, freqRange))
        % ��������� �����������
        assignin('base', Smn, reshape(evalin('base', Smn),[length, 1]));
        
        % ��� ���������� SmndB
        SmndB = SParameterVarNames{2, ik};  
        assignin('base', SmndB, 20 * log10(abs(evalin('base', Smn))));
        
        % ��� ���������� SmnAng
        SmnAng = SParameterVarNames{3, ik}; 
        if showAnglesInDegrees
            assignin('base', SmnAng, rad2deg(angle(evalin('base', Smn))));     
        else
            assignin('base', SmnAng, angle(evalin('base', Smn)));
        end
        
        if i==1
            % ������������� ������
            assignin('base', [SmndB 'Array'], evalin('base', SmndB));
            assignin('base', [SmnAng 'Array'], evalin('base', SmnAng));        
            
            % ������������� ������� ���
            if m == 1 && n == 1 
                VSWRArray = (1 + 10.^(0.1*S11dB))./(1 - 10.^(0.1*S11dB));
            end
        else
            % ���������� ������� �������� 
            assignin('base', [SmndB 'Array'], ...
                cat(2, evalin('base', [SmndB 'Array']), evalin('base', SmndB)));       
            assignin('base', [SmnAng 'Array'], ...
                cat(2, evalin('base', [SmnAng 'Array']), evalin('base', SmnAng)));
            
            % ���������� ������� �������� ��� 
            if m==1 && n==1
            VSWRArray = cat(2, VSWRArray, (1+10.^(0.1*S11dB))./(1-10.^(0.1*S11dB)));        
            end
        end
    % �������� ������ �������� (��� �������� � ������� ������������)
     clear(SParameterVarNames{:});
    end
end

% ������ ������������� ��
DNArray = zeros([length(1),width]);
for row = 1:length
  for col = 1:width
      DNArray(row,col) = S21dBArray(row,col) - max(S21dBArray(row,:));
  end
end

if saveTable
    % ���������� ��������
    for ik = 1:numel(saveList)

        % ��� ����� �������, �������� ������ ����������
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