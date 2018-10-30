clc
clear
close all

%%SOMETHING BROKEN IN WHILE LOOPS CHECKING IF MINDIST > 0, MOVING DOWN WHEN
%%SHOULDN'T?  something with datenum?  rounding error?  why???

%to test, run with check on or off
check = 0;

%subjlist = {'AG';'JM';'RN'};

subjlist = {'AG2';'RN2'};


for s=1:length(subjlist)
    subjID = subjlist{s};
    
    wd = fullfile('C:\Users\ppxrn\ExerciseAndAgeing\development\ergospect_testing\cosmed',subjID);
    
    cd(wd);
    %make processed/unproccesd
    sumfilepath = fullfile(wd,'processed','summary');
    mkdir(sumfilepath);
    rawfilepath = fullfile(wd,'processed','raw');
    mkdir(rawfilepath);
    
    files = dir('*CPET*');
    
    %make matlab float values of times to sample
%     sptimes = {'02:30'; '02:58'; '05:30'; '05:58'; '08:30'; '08:58'; '11:30'; '11:58'; '14:30';...
%         '14:58'; '17:30'; '17:58'; '20:30'; '20:58'; '23:30'; '23:58'; '26:30'; '26:58'};
    
%     sptimes = {'02:30'; '02:59'; '05:30'; '05:59'; '08:30'; '08:59'; '11:30'; '11:59'; '14:30';...
%         '14:59'; '17:30'; '17:59'; '20:30'; '20:59'; '23:30'; '23:59'; '26:30'; '26:59'};
    
    sptimes = {'02:30'; '03:00'; '05:30'; '06:00'; '08:30'; '09:00'; '11:30'; '12:00'; '14:30';...
        '15:00'; '17:30'; '18:00'; '20:30'; '21:00'; '23:30'; '24:00'; '26:30'; '27:00'};

%     sptimes = {'02:30.00'; '03:00.00'; '05:30.00'; '06:00.00'; '08:30.00'; '09:00.00'; '11:30.00'; '12:00.00'; '14:30.00';...
%         '15:00.00'; '17:30.00'; '18:00.00'; '20:30.00'; '21:00.00'; '23:30.00'; '24:00.00'; '26:30.00'; '27:00.00'};

    
    % sample times in matlab format
%     ntimes = datenum(sptimes,'MM:SS.FFF');
    ntimes = datenum(sptimes,'MM:SS');
    
    
    workload = [50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325];
    workload = transpose(workload);
    
    filename = files(1).name;
    sumfilename = fullfile(sumfilepath, strcat(filename(1:end-5),'_summary.xlsx'));
    
    %another loop for per subject, cycle through subject list
    %for s=1:lenght(subjlist)
    for f = 1:length(files)
        clear rs re Workload VE VCO2 RQ VO2kg HR perExcInt VO2excInt vo2relworkload
        clear num txt raw dtimes ddtimes dddtimes starttimes endtimes
        filename = files(f).name;
        
        [num,txt,raw] = xlsread(filename);
        
        % timestamps from real data
        dtimes(1:3,1) = 0;
        dtimes(4:length(num),1) = num(4:end,9);
        
%        ddtimes = datestr(dtimes, 'MM:SS.FFF');
        ddtimes = datestr(dtimes, 'MM:SS');

        %why do you need to do this?  why are dtimes numbers different?
%         dddtimes = datenum(ddtimes, 'MM:SS.FFF');
        dddtimes = datenum(ddtimes, 'MM:SS');
        
        %get last dddtime
        maxdtime = max(dddtimes);
        %find closest ntime to end dddtime
        [mindists, finalt] = min(abs(ntimes-maxdtime));
        %checks if final time index is odd (odd = start time), if so, use
        %nearest end time
        if mod(finalt,2)==1
            finalt = finalt-1;
        end
        finalw = (finalt/2)-1;
        
        %if filename has *conf* skip for now - need to modify to work
        if contains(filename, 'conf') == 1 || contains(filename, 'Conf') == 1
            %continue
            % could maybe get this out auto if read in corresponding ergospect
            % script?  would have to have good naming conventions to get it to
            % work
            n=1;
            h=1;
            w=1;
            confwkld = input(['What is the starting confirmation workload for ',subjID,'? (after 50) : ']);
            windex = find(workload==confwkld);
            %tindex = windex*2;
            for n = 1
                % find value in num(4:end, 10) closest to ntimes(n)
                [mindists, rs(h)] = min(abs(dddtimes-ntimes(n)));
                % rs = that value's row
                % find value in num(4:end, 10) closest to ntimes(n+1)
                [mindiste, re(h)] = min(abs(dddtimes-ntimes(n+1)));
                %checks to make sure endtime is not in next workload
                if check == 1
                    while mindiste > 0
                        disp(n);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                        re(h) = re(h)-1;
                        mindiste = dddtimes(re(h))-ntimes(n+1);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                    end
                end
                % re = that value's row
                VE(h,1) = mean(num(rs(h):re(h), 12));
                VCO2(h,1) = mean(num(rs(h):re(h), 15));
                RQ(h,1) = mean(num(rs(h):re(h), 16));
                VO2kg(h,1) = mean(num(rs(h):re(h),21));
                HR(h,1) = mean(num(rs(h):re(h), 23));
                starttimes{h,1} = ddtimes(rs(h),1:5);
                endtimes{h,1} = ddtimes(re(h),1:5);
                h=h+1;
                w=w+1;
            end
            for n = 3:2:finalt
                % find value in num(4:end, 10) closest to ntimes(n)
                [mindists, rs(h)] = min(abs(dddtimes-ntimes(n)));
                % rs = that value's row
                % find value in num(4:end, 10) closest to ntimes(n+1)
                [mindiste, re(h)] = min(abs(dddtimes-ntimes(n+1)));
                %checks to make sure endtime is not in next workload
                if check == 1
                    while mindiste > 0
                        disp(n);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                        re(h) = re(h)-1;
                        mindiste = dddtimes(re(h))-ntimes(n+1);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                    end
                end
                % re = that value's row
                VE(h,1) = mean(num(rs(h):re(h), 12));
                VCO2(h,1) = mean(num(rs(h):re(h), 15));
                RQ(h,1) = mean(num(rs(h):re(h), 16));
                VO2kg(h,1) = mean(num(rs(h):re(h),21));
                HR(h,1) = mean(num(rs(h):re(h), 23));
                starttimes{h,1} = ddtimes(rs(h),1:5);
                endtimes{h,1} = ddtimes(re(h),1:5);
                h=h+1;
                w=w+1;
                n=n+2;
            end
            
            Workload = vertcat(50, workload(windex:finalw+windex-1));
            T3 = table(starttimes,endtimes,Workload,VE,VCO2,RQ,VO2kg,HR);
            
            writetable(T3,sumfilename,'WriteRowNames',true,'Range','N1');
            
            fig = figure(s);
            % plot workload vs mean VO2kg and find line of best fit
            plot(Workload, VO2kg,'g*'); hold on;
            xlabel('Workload (Watts)');
            ylabel('VO2 (ml/min/kg)');
            title('with confirmation values');
            set(gca,'XTick',0:25:250);
            
            movefile(filename, rawfilepath);
            
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
                %checks to make sure endtime is not in next workload
                if check == 1
                    while mindiste > 0
                        disp(n);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                        re(h) = re(h)-1;
                        mindiste = dddtimes(re(h))-ntimes(n+1);
                        disp(mindiste);
                        disp(re(h));
                        disp(ddtimes(re(h),1:5));
                    end
                end
                % re = that value's row
                VE(h,1) = mean(num(rs(h):re(h), 12));
                VCO2(h,1) = mean(num(rs(h):re(h), 15));
                RQ(h,1) = mean(num(rs(h):re(h), 16));
                VO2kg(h,1) = mean(num(rs(h):re(h),21));
                HR(h,1) = mean(num(rs(h):re(h), 23));
                starttimes{h,1} = ddtimes(rs(h),1:5);
                endtimes{h,1} = ddtimes(re(h),1:5);
                h=h+1;
                w=w+1;
                n=n+2;
            end
            
            fig = figure(s);
            % plot workload vs mean VO2kg and find line of best fit
            plot(workload(1:h-1), VO2kg,'b*-'); hold on;
            lnfit = polyfit(workload(1:h-1,1), VO2kg, 1);
            slope = lnfit(1);
            intercept = lnfit(2);
            wksp = linspace(0,workload(h+2));
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
            
            T1 = table(starttimes,endtimes,Workload,VE,VCO2,RQ,VO2kg,HR);
            
            T2 = table(perExcInt,VO2excInt,vo2relworkload);
            
            writetable(T1,sumfilename,'WriteRowNames',true);
            writetable(T2,sumfilename,'Range','J1');
            xlswritefig(fig(f),sumfilename,'Sheet1','A12');
            movefile(filename, rawfilepath);
        end
        
    end
    
    xlswritefig(fig,sumfilename,'Sheet1','N12');
    
end

