%% Load data
M = readcell('InterviewTracker.xlsx','Sheet','MPP Dom Waitlist 2', 'NumHeaderLines',1);
filename_template = 'Template_Coversheet Domestic.docx';

%% Make folders
N = size(M, 1);
fullname_ID = cell(N,1);
for i = 1 : N
    fullname_ID{i} = [M{i,1} ' ' upper(M{i,2}) ' ' num2str(M{i,3})]; % Create First LAST STUDENT ID
    mkdir(fullname_ID{i});                                           % Make new folder for student
    
end

%% Move templates
fileList_word = dir('*.docx');
for i = 1 : N
    for ii = 1 : size(fileList_word, 1)
            
            filename_word = fileList_word(ii).name;
            % Compares first and last names to match - also replaces the last name spaces with '-'
            if strcmpi(erase(filename_word, '.docx'), fullname_ID{i}) 
                movefile(filename_word , fullname_ID{i});
            end
    end
end

%% Referee reports
fileList_pdf = dir('*.pdf');
num_referee = zeros(N,1);
missing = {};

for i = 1 : N
    try 
        
        for ii = 1 : size(fileList_pdf, 1)
            
            filename_referee = fileList_pdf(ii).name;
            name = extractBetween(filename_referee ,'_', '_');

            % Compares first and last names to match - also replaces the last name spaces with '-'
            if strcmpi(name{1}, strrep([M{i,1}, '-', M{i,2}], ' ', '-')) 
                num_referee(i) = num_referee(i) + 1;
                filename_rename_referee = ['Referee ', num2str(num_referee(i)), ' - ' M{i,1} ' ' M{i,2} ' ' num2str(M{i,3}) '.pdf'];
                movefile(filename_referee , filename_rename_referee);
                movefile(filename_rename_referee , fullname_ID{i});
            end
        end
        
    catch E
        missing{end} = filename_referee;
        error(E);
    end
end

%writecell([num2cell(num_referee) M], 'Number of Referees.xlsx');
return

%% Move transcripts

