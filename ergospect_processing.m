%ergospect processing

cd C:\Users\ppxrn\ExerciseAndAgeing\development\ergospect_testing\cosmed

clear  
close all
clc
%set subjid

subjlist = {'RN2';'AG2'};

for s = 1%:length(subjlist)
    subjid = subjlist{s};

cd(subjid)

files = dir('*Measurement*');

for f = 3%:length(files)
    
    % conf if loop
    %        if contains(filename, 'conf') == 1 || contains(filename, 'Conf') == 1

    
    
% read in csv file
filename = files(f).name;
sumfilename = strcat(filename(1:end-4),'_summary.xlsx');

%num = dlmread(filename,';');
%num = readErgo(filename,14,inf);

[num,txt,raw] = xlsread(filename);

%sometimes 12, sometimes 13??
rawtime = num(13:end,2);
force1 = num(13:end,4);
force2 = num(13:end,5);
frequency = num(13:end,6);
power = num(13:end,7);
way = num(13:end,9);
work = num(13:end,10);

sptimes = {'02:30'; '03:00'; '05:30'; '06:00'; '08:30'; '09:00'; '11:30'; '12:00'; '14:30';...
    '15:00'; '17:30'; '18:00'; '20:30'; '21:00'; '23:30'; '24:00'; '26:30'; '27:00'};

ntimes = datenum(sptimes,'MM:SS');


%convert timestamps into time elapsed 
% will break if are Nans in data... 
eltime = rawtime-rawtime(1);
streltime = datestr(eltime,'MM:SS');
reltime = datenum(streltime,'MM:SS');

maxeltime = max(reltime);
%find closest ntime to end dddtime
[mindists, finalt] = min(abs(ntimes-maxeltime));

h =1;

for n=1:2:finalt
    [mindists, rs(h)] = min(abs(reltime-ntimes(n)));
% rs = that value's row
% find value in num(4:end, 10) closest to ntimes(n+1)
    [mindiste, re(h)] = min(abs(reltime-ntimes(n+1)));

    mfreq(h,1) = mean(frequency(rs(h):re(h)));
    mpower(h,1) = mean(power(rs(h):re(h)));
    mwork(h,1) = mean(work(rs(h):re(h)));
    peaks{h} = findpeaks(way(rs(h):re(h)));
    mway(h,1) = mean(peaks{h});
    
    starttimes{h,1} = reltime(rs(h));
    endtimes{h,1} = reltime(re(h));
    
    sstarttimes{h,1} = streltime(rs(h),1:5);
    sendtimes{h,1} = streltime(re(h),1:5);
    h=h+1;
    n=n+2;
end

mtendtimes = cell2mat(endtimes);

% plot freq vs endtimes (endtimes in matformat now, need to
% convert to something legible)
figf = figure(1);
plot(mtendtimes,mfreq, 'b*'); hold on;
fline = refline([0 70]);
fline.Color = 'r';
xlabel('Time');
ylabel('mean frequency');
ylim([50 80]);


figp = figure(2);
plot(mtendtimes,mpower, 'b*');  hold on;
plot(mtendtimes,[50 225]); % need to edit to read workload
xlabel('Time');
ylabel('mean power');


figw = figure(3);
plot(mtendtimes,mwork, 'b*'); hold on;
xlabel('Time');
ylabel('mean work');

figwy = figure(4);
plot(mtendtimes,mway, 'k*');
xlabel('Time');
ylabel('mean way');


T1 = table(sstarttimes,sendtimes,mfreq,mpower,mwork);

writetable(T1,sumfilename);

xlswritefig(figf,sumfilename,'Sheet1','A12');
xlswritefig(figp,sumfilename,'Sheet1','J12');
xlswritefig(figw,sumfilename,'Sheet1','T12');


%movefile(filename, rawfilepath);


end
end

%find peaks of force and get average force (per foot - 1 and 2 -
...average together) for 3 min and 30 sec periods
    