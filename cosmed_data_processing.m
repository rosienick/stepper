clc
clear
close all

%%WON'T HAVE RIGHT WORKLOAD VALUES FOR CONFIRMATION RUNS

wd = 'C:\Users\ppxrn\ExerciseAndAgeing\development\ergospect_testing\cosmed';

cd(wd);
%make processed/unproccesd
sumfilepath = fullfile(wd,'processed','summary');
rawfilepath = fullfile(wd,'processed','raw');

files = dir('*CPET*');

%make matlab float values of times to sample
sptimes = {'02:30'; '02:58'; '05:30'; '05:58'; '08:30'; '08:58'; '11:30'; '11:58'; '14:30';...
    '14:58'; '17:30'; '17:58'; '20:30'; '20:58'; '23:30'; '23:58'; '26:30'; '26:58'};

% sample times in matlab format
ntimes = datenum(sptimes,'MM:SS');

workload = [50, 75, 100, 125, 150, 175, 200, 225, 250];
workload = transpose(workload);

for f = 1:length(files)
    clear rs re Workload VE VCO2 RQ VO2kg HR perExcInt VO2excInt vo2relworkload
    clear num txt raw dtimes ddtimes dddtimes
    filename = files(f).name;
    sumfilename = fullfile(sumfilepath, strcat(filename(1:end-5),'_summary.xlsx'));

    [num,txt,raw] = xlsread(filename);
    
    % timestamps from real data
    dtimes(1:3,1) = 0;
    dtimes(4:length(num),1) = num(4:end,9);
    
    ddtimes = datestr(dtimes, 'MM:SS');
    %why do you need to do this?  why are dtimes numbers different?
    dddtimes = datenum(ddtimes, 'MM:SS');
    
    %get last dddtime
    maxdtime = max(dddtimes);
    %find closest ntime to end dddtime
    [mindists, finalt] = min(abs(ntimes-maxdtime));
    
    %if filename has *conf* skip for now - need to modify to work
    if contains(filename, 'conf') == 1 || contains(filename, 'Conf') == 1
        continue
        % could maybe get this out auto if read in corresponding ergospect
        % script?  would have to have good naming conventions to get it to
        % work
        %confwkld = input('What is the starting confirmation workload? (after 50) : ');
        %windex = find(workload==confwkld);
        % for n = 1,, workload(windex:finalt)
        
        %T3 = table()
        %writetable(T3,sumfilename,'WriteRowNames',true,'Range','M1');
        
    else
        
        %last workload
        %lwl = input('what is the final workload you want measures for? : ');
        n=1;
        h=1;
        w=1;
        %while lwl >= workload(w)
        %times to collect data points - nearest value
        % in raw or num - 4:end,10
        for n = 1:2:finalt 
            % find value in num(4:end, 10) closest to ntimes(n)
            [mindists, rs(h)] = min(abs(dddtimes-ntimes(n)));
            % rs = that value's row
            % find value in num(4:end, 10) closest to ntimes(n+1)
            [mindiste, re(h)] = min(abs(dddtimes-ntimes(n+1)));
            % re = that value's row
            VE(h,1) = mean(num(rs(h):re(h), 12));
            VCO2(h,1) = mean(num(rs(h):re(h), 15));
            RQ(h,1) = mean(num(rs(h):re(h), 16));
            VO2kg(h,1) = mean(num(rs(h):re(h),21));
            HR(h,1) = mean(num(rs(h):re(h), 23));
            starttimes{h} = ddtimes(rs(h),1:5);
            endtimes{h} = ddtimes(re(h),1:5);
            h=h+1;
            w=w+1;
            n=n+2;
        end
        
        fig(f) = figure(f);
        % plot workload vs mean VO2kg and find line of best fit
        plot(workload(1:h-1), VO2kg,'b*-'); hold on;
        lnfit = polyfit(workload(1:h-1,1), VO2kg, 1);
        slope = lnfit(1);
        intercept = lnfit(2);
        wksp = linspace(0,workload(h-1));
        funln = polyval(lnfit,wksp);
        plot(wksp,funln, 'r--');
        xlabel('Workload (Watts)');
        ylabel('VO2 (ml/min/kg)');
        title(strcat('y =',num2str(slope),'x + ',num2str(intercept)));
        set(gca,'XTick',0:25:250);
        
        
        %relative exercise intensity
        
        VO2peak = max(VO2kg);
        % percent of VO2 peak - 30, 50, 70
        perExcInt = [.3 .5 .7];
        perExcInt = transpose(perExcInt);
        
        for p=1:length(perExcInt)
            VO2excInt(p,1) = perExcInt(p,1)*VO2peak;
            % workload = (VO2 at % - line intercept) / slope
            vo2relworkload(p,1) =  (VO2excInt(p,1) - intercept) / slope;
        end
        
        % output tables with data
        Workload = workload(1:h-1,1);
        
        T1 = table(Workload,VE,VCO2,RQ,VO2kg,HR,'RowNames',endtimes);
        T2 = table(perExcInt,VO2excInt,vo2relworkload);
                
        writetable(T1,sumfilename,'WriteRowNames',true);
        writetable(T2,sumfilename,'Range','I1');
        xlswritefig(fig(f),sumfilename,'Sheet1','A12');
        
        movefile(filename, rawfilepath);
    end
end


