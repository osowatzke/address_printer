function s = loadAddresses
    C = readcell('Addresses.xlsx');
    C = C(3:end,:);
    s = {};
    newEntry = true;
    for i = 1:length(C)
        if any(ismissing(C{i,1})) && any(ismissing(C{i,2}))
            newEntry = true;
        else
            if newEntry
                s{end+1} = struct('name',[],'addr',[]);
                s{end}.name = {};
                s{end}.addr = {};
            end
            newEntry = false;
            if ~any(ismissing(C{i,1}))
                s{end}.name{end+1} = C{i,1};
            end
            if ~any(ismissing(C{i,2}))
                s{end}.addr{end+1} = C{i,2};
            end
        end
    end
end