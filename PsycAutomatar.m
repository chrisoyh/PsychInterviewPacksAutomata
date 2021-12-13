% REMINDERS: 
% * Have all your Referee Reports, Coversheet templates and Excel file containing all the GivenName, 
% FamilyName and Student ID in the same folder as this script
% * Ensure no students have the same GivenName and FamilyName
% * Coversheets must be .docx, Referee Reports must be .pdf, ensure that no other files of these 2 
% formats are present in the folder

%% Load data (always run this section first)
M = readcell('InterviewTracker.xlsx','Sheet','MPP Dom Waitlist 2', 'NumHeaderLines',1); % Reads GivenName, LastName and StudentID from Excel

%% Make folders
N = size(M, 1);
fullname_ID = cell(N,1);
for i = 1 : N
    fullname_ID{i} = [M{i,1} ' ' upper(M{i,2}) ' ' num2str(M{i,3})]; % Create cell array with {'GivenName' 'LASTNAME' 'StudentID'}
    mkdir(fullname_ID{i});                                           % Make new applicant folder
end

%% Move Coversheets into folders
fileList_word = dir('*.docx'); % Reads coversheet filenames
for i = 1 : N
    for ii = 1 : size(fileList_word, 1)
            
            filename_word = fileList_word(ii).name;
            % Compares first and last names to match - also replaces the last name spaces with '-'
            if strcmpi(erase(filename_word, '.docx'), fullname_ID{i}) 
                movefile(filename_word , fullname_ID{i}); % moves coversheet to the applicants folder
            end
    end
end

%% Referee reports
fileList_pdf = dir('*.pdf'); % Reads referee reports filenames
num_referee = zeros(N,1); % Number of referee reports per student 
missing = {}; % Check this variable in case there are students that are missed from errors

for i = 1 : N
    try 
        
        for ii = 1 : size(fileList_pdf, 1)
            
            filename_referee = fileList_pdf(ii).name;
            name = extractBetween(filename_referee ,'_', '_');

            % Compares first and last names to match - also replaces the last name spaces with '-'
            if strcmpi(name{1}, strrep([M{i,1}, '-', M{i,2}], ' ', '-')) 
                num_referee(i) = num_referee(i) + 1; % Counts the number of referee reports per student
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

% Checks for students that have less than 2 referee reports
idx_less_than_2_reports = num_referee < 2;
less_than_2_reports = M{idx_less_than_2_reports,3};


return

