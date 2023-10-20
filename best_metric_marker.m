function best_metric_marker(data, num_metrics, varargin)

% Input
%   data: the metrics.
%   num_metrics: the number of metrics.
%   varargin: Other arguments, including
%        'precision': if set 'precision' to N, the code will round data to
%                     N digits to the right of the decimal point. (Default: 4)
%        'optval': it refers to the direction of the best metrics. Set is
%                  as 'max' if the higher metric corresponds to better
%                  results; and 'min' otherwise. (Default: 'max')
%        'highlight': it determines which kind of marker is applied to the
%                     the metric. Support 'bold', 'underline', 'italic',
%                     'color'. For example, set the best metric as bold,
%                     and the 2nd best metric as red, and then 'highlight'
%                     can be set as
%                         highlight.key = {'bold', 'color'};
%                         highlight.value = {true, [255,0,0]};
%        'border_mode': it controls the border of cells. Support 'top',
%                       'bottom', 'mid'.  (Default: {})
%        'filename': the path to write the Excel file.
%                    (Default: 'best_metric_time.xlsx')
%
%   Usage:
%      Suppose you have obtain the metrics of n remote sensing image
%      processing algorithms on different datasets, and the metric is
%      organized as follows:
%                      |  dataset1     dataset2
%      -------------------------------------------
%      PSNR  of alg.1  |      p11           p12
%      SSIM  of alg.1  |      s11           s12
%      ERGAS of alg.1  |      e11           e12
%      SAM   of alg.1  |      a11           a12
%      -------------------------------------------
%      PSNR  of alg.2  |      p21           p22
%      SSIM  of alg.2  |      s21           s22
%      ERGAS of alg.2  |      e21           e22
%      SAM   of alg.2  |      a21           a22
%      -------------------------------------------
%             .                .             .
%             .                .             .
%             .                .             .
%      -------------------------------------------
%      PSNR  of alg.n  |      pn1           pn2
%      SSIM  of alg.n  |      sn1           sn2
%      ERGAS of alg.n  |      en1           en2
%      SAM   of alg.n  |      an1           an2
%      -------------------------------------------
%      where higher PSNR and SSIM lead to better results, and lower ERGAS
%      and SAM lead to better results. We want to round SSIM values as 3
%      digits, and round others as 2 digits. And mark the 1st, 2nd and
%      3rd best metrics as 'bold', 'color', 'italic'. The code should
%      be
%
%      num_metrics = 4;
%      precision = [2,3,2,2];
%      optval = {'max', 'max', 'min', 'min'};
%      highlight.key = {'bold', 'color', 'italic'};
%      highlight.value = {true, [255,0,0], true};
%      border_mode = {'top', 'bottom', 'mid'};
%      filename = 'my_metrics.xlsx';
%      best_metric_marker(data,  ...
%                         num_metrics,  ...
%                         'precision', precision, ...
%                         'optval', optval, ...
%                         'border_mode', border_mode, ...
%                         'highlight', highlight, ...
%                         'filename', filename);
% 
% Copyright (c) 2023 Shuang Xu
% Email: xu.s@outlook.com; xs@nwpu.edu.cn


% =========================
% parse the input arguments
% =========================
highlight_mode.key = {'bold', 'underline', 'italic'};
highlight_mode.value = {true, true, true};
filename = strcat('best_metric_',datestr(now), '.xlsx'); 
filename = replace(filename, ':', '-');
savepath = pwd; % current path 
[precision, optval, highlight, border_mode, filename] = ...
    process_options(varargin, ...
    'precision',     4*ones(1,num_metrics), ...
    'optval',        repmat('max', 1, num_metrics), ...
    'highlight',     highlight_mode, ...
    'border_mode',   {}, ...
    'filename',      fullfile(savepath,filename));

% =========================
% Main code
% =========================

% Preprocess data precision
% 预处理：调整数据保留的小数点位数
for i = 1:size(data,1)/num_metrics
    for j = 0:num_metrics-1
        data(num_metrics*i-num_metrics+1+j,:) = round(data(num_metrics*i-num_metrics+1+j,:),precision(j+1));
    end
end

% Define letters
% 定义字母
letters = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'};

% Create an Excel object.
% 创建Excel
Excel=actxserver('Excel.application');

% Add a workbook.
% 添加workbook
Excel.Visible = 1;
Workbook = Excel.Workbooks.Add;

% Make the first sheet active.
% 激活一个sheet
ExcelActiveWorkbooks = get(Excel, 'ActiveWorkBook');
eSheet = Excel.ActiveWorkbook.Sheets;
eSheet1 = Item(eSheet, 1);

% Put MATLAB data into the worksheet.
% 写入单元格内容
eSheet1.Range(['A1:',letters{size(data,2)},num2str(size(data,1))]).Value = data;

% Configure the format
% 设置单元格格式
eSheet1.Range(['A1:',letters{size(data,2)},num2str(size(data,1))]).Font.Size = 11; % set fontsize 设置单元格字体大小
eSheet1.Range(['A1:',letters{size(data,2)},num2str(size(data,1))]).Font.name='Arial'; % set font family设置单元格字体
eSheet1.Range(['A1:',letters{size(data,2)},num2str(size(data,1))]).HorizontalAlignment=-4108;  % set Horizontal Alignment 水平居中
eSheet1.Range(['A1:',letters{size(data,2)},num2str(size(data,1))]).VerticalAlignment=-4108; % set Vertical Alignment 垂直居中


% Mark the best metrics
% 标记最优指标
total_line = size(data,1);
for i = 1:num_metrics
    linenum = i:num_metrics:total_line;
    temp_metric = data(linenum,:);
    if strcmp(optval{i}, 'max')
        [~,ind] = sort(temp_metric,'descend');
    elseif strcmp(optval{i}, 'min')
        [~,ind] = sort(temp_metric,'ascend');
    end
    ind = linenum(ind);

    for k = 1:size(ind,2)
        for j = 1:length(highlight.key)
            coor = [letters{k},num2str(ind(j,k))];
            set_highlight(eSheet1, coor, highlight.key{j}, highlight.value{j})
        end
    end

end


% Set borders
% 设置边框
if ismember('top', border_mode)
    eSheet1.Range(['A','1',':',letters{size(data,2)},'1']).Borders.Item(3).Weight=2;%设定表格的上边框为线段加粗
end
if ismember('mid', border_mode)
    for i = num_metrics:num_metrics:size(data,1)-1
        eSheet1.Range(['A',num2str(i),':',letters{size(data,2)},num2str(i)]).Borders.Item(4).Weight=2;%设定表格的下边框为线段加粗
    end
end
if ismember('bottom', border_mode)
    eSheet1.Range(['A',num2str(size(data,1)),':',letters{size(data,2)},num2str(size(data,1))]).Borders.Item(4).Weight=2;%设定表格的下边框为线段加粗
end

% Save the workbook in a file.
% 保存Excel文件
SaveAs(Workbook,filename)

% Close the workbook.
% 关闭
Close(Workbook)

% Quit the Excel program and delete the server object.
% 退出&删除
Quit(Excel)
delete(Excel)

end







% Subfunction to highlight best metrics
function set_highlight(eSheet1, coor, mode, mode_value)
if strcmpi(mode, 'bold')
    eSheet1.Range(coor).Font.Bold = true;
elseif strcmpi(mode, 'underline')
    eSheet1.Range(coor).Font.Underline = true;
elseif strcmpi(mode, 'italic')
    eSheet1.Range(coor).Font.Italic = true;
elseif strcmpi(mode, 'color')
    % see the anwser from Walter Roberson at
    % https://www.mathworks.com/matlabcentral/answers/3352-how-to-set-excel-cell-color-to-red-from-matlab
    C = double(mode_value(1)) * 256^0 + double(mode_value(2)) * 256^1 + double(mode_value(3)) * 256^2;
    eSheet1.Range(coor).Font.Color = C;
end
end % end prepareArgs


function [varargout] = process_options(args, varargin)

args = prepareArgs(args); % added to support structured arguments
% Check the number of input arguments
n = length(varargin);
if (mod(n, 2))
    error('Each option must be a string/value pair.');
end

% Check the number of supplied output arguments
if (nargout < (n / 2))
    error('Insufficient number of output arguments given');
elseif (nargout == (n / 2))
    warn = 1;
    nout = n / 2;
else
    warn = 0;
    nout = n / 2 + 1;
end

% Set outputs to be defaults
varargout = cell(1, nout);
for i=2:2:n
    varargout{i/2} = varargin{i};
end

% Now process all arguments
nunused = 0;
for i=1:2:length(args)
    found = 0;
    for j=1:2:n
        if strcmpi(args{i}, varargin{j}) || strcmpi(args{i}(2:end),varargin{j})
            varargout{(j + 1)/2} = args{i + 1};
            found = 1;
            break;
        end
    end
    if (~found)
        if (warn)
            warning(sprintf('Option ''%s'' not used.', args{i}));
            args{i}
        else
            nunused = nunused + 1;
            unused{2 * nunused - 1} = args{i};
            unused{2 * nunused} = args{i + 1};
        end
    end
end

% Assign the unused arguments
if (~warn)
    if (nunused)
        varargout{nout} = unused;
    else
        varargout{nout} = cell(0);
    end
end

end % end process_options

function out = prepareArgs(args)
% Convert a struct into a name/value cell array for use by process_options
%
% Prepare varargin args for process_options by converting a struct in args{1}
% into a name/value pair cell array. If args{1} is not a struct, args
% is left unchanged.
% Example:
% opts.maxIter = 100;
% opts.verbose = true;
% foo(opts)
%
% This is equivalent to calling
% foo('maxiter', 100, 'verbose', true)

% This file is from pmtk3.googlecode.com


if isstruct(args)
    out = interweave(fieldnames(args), struct2cell(args));
elseif ~isempty(args) && isstruct(args{1})
    out = interweave(fieldnames(args{1}), struct2cell(args{1}));
else
    out = args;
end

end % end prepareArgs
