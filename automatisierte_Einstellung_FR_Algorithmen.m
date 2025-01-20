%Model ermöglicht eine Anpassung der Verbindungen mit einer reinen
%Anpassung des Ausgangspunktes um den Fehlerort anzupassen
%Modell erweitert um einzelne Einstellmöglichkeiten der Algorithmen zu
%testen

% Prompt the user to select or create a folder for saving scenario data and boxplots
folderPath = uigetdir('', 'Select or Create a Folder to Save Scenario Data and Boxplots');

% Check if the user canceled the folder selection
if isequal(folderPath, 0)
    error('No folder selected. Please select a valid folder.');
else
    fprintf('Selected folder: %s\n', folderPath);
end

% Create a subfolder for saving results within the selected folder
resultFolderName = 'Simulation_Results';
resultFolderPath = fullfile(folderPath, resultFolderName);

% Create the folder if it does not exist
if ~exist(resultFolderPath, 'dir')
    mkdir(resultFolderPath);
    fprintf('Created result folder: %s\n', resultFolderPath);
else
    fprintf('Result folder already exists: %s\n', resultFolderPath);
end
% Use inputdlg to prompt for the simulation stop time
stopTimePrompt = {'Enter the stop time for the simulation (in seconds):'};
dlgTitle = 'Simulation Stop Time';
numLines = 1;
defaultAnswer = {'1'}; % Default stop time in seconds
answer = inputdlg(stopTimePrompt, dlgTitle, numLines, defaultAnswer);
stopTime = answer{1};
% Use uigetfile to select the Excel file
[excelFileName, excelFilePath] = uigetfile('*.xlsx', 'Select the Excel file with scenario data');

stationTime= 200; % Stationäres Zeitfenster welcher zur Auswertung des Winkels phi verwendet wird



% Read the data from the Excel file
scenarioTable = readtable(excelFileName);
measurmentPointTable= readtable("measuringPointsBenchmark.xlsx");

% Define a fixed "to" connection
fixedToConnection1 = 'Messen_F'; %Fixer Punkt Fehler
fixedToConnection2 = 'VoltageScopeVF';
fixedToConnection3 = 'CurrentScopeVF';
fixedToConnection4 = 'VoltageScopeHF';
fixedToConnection5 = 'CurrentScopeHF';
fixedToConnection6 = 'VoltageScopeGF';
fixedToConnection7 = 'CurrentScopeGF';
fixedToConnection8 = 'VoltageScopeGAF';
fixedToConnection9 = 'CurrentScopeGAF';
% fixedToConnection10 = 'threeToOnePhase';

% Initialize an empty structure for scenarios
scenarios = struct();

% Extract unique scenario names
uniqueScenarios = unique(scenarioTable.ScenarioName,'stable');

% Number of unique scenarios
numScenarios = length(uniqueScenarios);

% Loop through each unique scenario to populate the scenarios structure
for i = 1:height(uniqueScenarios)
    scenarioName = scenarioTable.ScenarioName{i};
    scenarioName = uniqueScenarios{i};
    scenarioRows = strcmp(scenarioTable.ScenarioName, scenarioName);
    connections = scenarioTable(scenarioRows, :);
    faultedLine = scenarioTable.FromBusbar{i};
    measName = uniqueScenarios{i};
    fromBusbar = scenarioTable.FromBusbar{i};
    additionalConnectionsRows = strcmp(measurmentPointTable.faultPoint, fromBusbar);
    additionalConnections = measurmentPointTable(additionalConnectionsRows, :);
    measconnections = measurmentPointTable(additionalConnectionsRows, :);
    measurementvVF = measurmentPointTable(additionalConnectionsRows, :);
    measurementiVF = measurmentPointTable(additionalConnectionsRows, :);
    measurementvHF = measurmentPointTable(additionalConnectionsRows, :);
    measurementiHF = measurmentPointTable(additionalConnectionsRows, :);
    measurementvGF = measurmentPointTable(additionalConnectionsRows, :);
    measurementiGF = measurmentPointTable(additionalConnectionsRows, :);
    measurementvGAF = measurmentPointTable(additionalConnectionsRows, :);
    measurementiGAF = measurmentPointTable(additionalConnectionsRows, :);
    % measurementFaultTime = measurmentPointTable(additionalConnectionsRows, :);
        
    scenarios(i).name = scenarioName;
    for j = 1:height(connections)
        scenarios(i).connections(j).from = connections.FromBusbar{j}; %Fehlerort
        scenarios(i).connections(j).to = fixedToConnection1;
        scenarios(i).faultResistance = connections.FaultResistance(j); %Fehlerwiderstand
        scenarios(i).faultedPhase = connections.FaultedPhase{j};%Fehlerhafte Phase
        scenarios(i).faultTime= connections.FaultTime(j);
        scenarios(i).measurementvVF(j).to = fixedToConnection2;
        scenarios(i).measurementiVF(j).to = fixedToConnection3;
        scenarios(i).measurementvHF(j).to = fixedToConnection4;
        scenarios(i).measurementiHF(j).to = fixedToConnection5;
        scenarios(i).measurementvGF(j).to = fixedToConnection6;
        scenarios(i).measurementiGF(j).to = fixedToConnection7;
        scenarios(i).measurementvGAF(j).to = fixedToConnection8;
        scenarios(i).measurementiGAF(j).to = fixedToConnection9;
        % scenarios(i).measurementFaultTime(j).to = fixedToConnection10;
    end
    
  
    for k = 1:height(additionalConnections)
            % Add additional connections to the scenario structure
            additionalConnectionIndex = numel(scenarios(i).connections);
            scenarios(i).measurementvVF(additionalConnectionIndex).from = measurementvVF.measurementvVF{k};
            scenarios(i).measurementiVF(additionalConnectionIndex).from = measurementiVF.measurementiVF{k};%Messpunkt VF
            scenarios(i).measurementvHF(additionalConnectionIndex).from = measurementvHF.measurementvHF{k};
            scenarios(i).measurementiHF(additionalConnectionIndex).from = measurementiHF.measurementiHF{k};%Messpunkt HF
            scenarios(i).measurementvGF(additionalConnectionIndex).from = measurementvGF.measurementvGF{k};
            scenarios(i).measurementiGF(additionalConnectionIndex).from = measurementiGF.measurementiGF{k};%Messpunkt GF
            scenarios(i).measurementvGAF(additionalConnectionIndex).from = measurementvGAF.measurementvGAF{k};
            scenarios(i).measurementiGAF(additionalConnectionIndex).from = measurementiGAF.measurementiGAF{k};%Messpunkt GAF
            % scenarios(i).measurementFaultTime(additionalConnectionIndex).from = measurementFaultTime.measurementFaultTime{k};%Messpunkt Fehlerzeitpunkt
        end

end

% Initialize a structure to hold all scenario data
allScenarioData = struct();

% Open the base model

baseModel = 'Benchmarknetz_Masterarbeit1810';

% Loop through each scenario
for i = 1:numel(scenarios)
    scenario = scenarios(i);
    load_system(baseModel);

    % Use a unique identifier for the model name
    newModel = [baseModel '_S' num2str(i)];
    
    % Check if the new model already exists and close it if it does
    if bdIsLoaded(newModel)
        close_system(newModel, 0);
    end
    
    % Try to save the base model to a new model
    try
        save_system(baseModel);
        fprintf('Model saved as: %s\n', baseModel);
        
    catch ME
        error('Failed to save new model: %s. Error: %s', newModel, ME.message);
    end
    
    % Try to open the new model
    try
        open_system(baseModel);
        fprintf('Model opened: %s\n', baseModel);
    catch ME
        error('Failed to open new model: %s. Error: %s', newModel, ME.message);
    end
    
    % Add new connections based on the scenario for fault parameters
    for j = 1:numel(scenario.connections)
        connection = scenario.connections(j);
        
        % Define phase suffixes
        phases1 = {'1', '2', '3'};
        phases2 = {'1', '2', '3'};

        for k = 1:numel(phases1)
            phase1 = phases1{k};
            phase2 = phases2{k};

            fromPhase = [connection.from '/RConn ' phase1];
            toPhase = [connection.to '/RConn ' phase2];
            
            try
                add_line(baseModel, fromPhase, toPhase);
            catch ME
                error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
            end
        end
    end

% VF: Add new connections for measuring points in relation to the fault 
if isfield(scenario, 'measurementvVF')
for additionalConnectionIndex = 1:numel(scenario.measurementvVF)
        meas1 =  scenario.measurementvVF(additionalConnectionIndex);
        meas2 =  scenario.measurementiVF(additionalConnectionIndex);
            fromPhaseV = [meas1.from '/1 '];
            toPhaseV = [meas1.to '/1 '];
            fromPhaseI = [meas1.from '/2 '];
            toPhaseI = [meas2.to '/1 '];
            try
                add_line(baseModel, fromPhaseV, toPhaseV);
                add_line(baseModel, fromPhaseI, toPhaseI);
            catch ME
                error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
            end
       end
end

% HF: Add new connections for measuring points in relation to the fault 
if isfield(scenario, 'measurementvHF')
for additionalConnectionIndex = 1:numel(scenario.measurementvHF)
        meas3 = scenario.measurementvHF(additionalConnectionIndex);
        meas4 = scenario.measurementiHF(additionalConnectionIndex);
            fromPhaseV = [meas3.from '/1 '];
            toPhaseV = [meas3.to '/1 '];
            fromPhaseI = [meas3.from '/2 '];
            toPhaseI = [meas4.to '/1 '];
            try
                add_line(baseModel, fromPhaseV, toPhaseV);
                add_line(baseModel, fromPhaseI, toPhaseI);
            catch ME
                error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
            end
      end
end
% GF: Add new connections for measuring points in relation to the fault 
if isfield(scenario, 'measurementvGF')    
    for k = 1:numel(scenario.measurementvGF)
        meas5 = scenario.measurementvGF(additionalConnectionIndex);
        meas6 = scenario.measurementiGF(additionalConnectionIndex);
            fromPhaseV = [meas5.from '/1 '];
            toPhaseV = [meas5.to '/1 '];
            fromPhaseI = [meas5.from '/2 '];
            toPhaseI = [meas6.to '/1 '];
            try
                add_line(baseModel, fromPhaseV, toPhaseV);
                add_line(baseModel, fromPhaseI, toPhaseI);
            catch ME
                error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
            end
    end
end
 % GAF: Add new connections for measuring points in relation to the fault 
   if isfield(scenario, 'measurementvGAF')
    for k = 1:numel(scenario.measurementvGAF)
        meas7 = scenario.measurementvGAF(additionalConnectionIndex);
        meas8 = scenario.measurementiGAF(additionalConnectionIndex);
            fromPhaseV = [meas7.from '/1 '];
            toPhaseV = [meas7.to '/1 '];
            fromPhaseI = [meas7.from '/2 '];
            toPhaseI = [meas8.to '/1 '];
            try
                add_line(baseModel, fromPhaseV, toPhaseV);
                add_line(baseModel, fromPhaseI, toPhaseI);
            catch ME
                error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
            end
    end    
   end

% % FehlerZeit: Add new connections for measuring points in relation to the fault 
%    if isfield(scenario, 'measurementFaultTime')
%     for k = 1:numel(scenario.measurementFaultTime)
%         meas9 = scenario.measurementFaultTime(additionalConnectionIndex);
%             fromPhase = [meas9.from '/1 '];
%             toPhase = [meas9.to '/1 '];
%             try
%                 add_line(baseModel, fromPhase, toPhase);
%             catch ME
%                 error('Failed to add line from %s to %s in model %s. Error: %s', fromPhase, toPhase, newModel, ME.message);
%             end
%     end    
%    end

    % Set up simulation parameters
    set_param(baseModel, 'SimulationMode', 'normal');
    
    % Set fault block parameters
    faultBlockPath = [baseModel '/Three-Phase Fault1'];
    stepBlockPath = [baseModel '/FaultTimeStep'];
    faultedPhase= scenario.faultedPhase;
    set_param(faultBlockPath, 'FaultResistance', num2str(scenario.faultResistance));
    set_param(stepBlockPath, 'Time', num2str(scenario.faultTime));
    blockParams = get_param(faultBlockPath, 'InitialStates');
    %set_param(faultBlockPath, 'InitialStates','0');
    
    

    set_param(faultBlockPath, 'FaultA', 'off');
    set_param(faultBlockPath, 'FaultB', 'off');
    set_param(faultBlockPath, 'FaultC', 'off');
    set_param(faultBlockPath, 'GroundFault', 'off');

    % Input char value
phaseStr = faultedPhase; % Alternatively, 'L1-N' if it's a char

% Map string to a numeric value
switch phaseStr
    case "L1-N"
        faultedPhaseNumber1 = 1; % Assign a numeric value for "L1-N"
    case "L2-N"
        faultedPhaseNumber1 = 2;
    case "L3-N"
        faultedPhaseNumber1 = 3;
    otherwise
        error("Unknown input string.");
end

% Convert to desired numeric type (e.g., int32)
faultedPhaseNumber= int32(faultedPhaseNumber1);


    switch scenario.faultedPhase
        case 'L1-N'
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L2-N'
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L3-N'
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L1L2'
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'GroundFault', 'off');
        case 'L1L2-N'
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L2L3-N'
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L2L3'
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'GroundFault', 'off');
        case 'L3L1-N'
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L3L1'
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'GroundFault', 'off');
        case 'L1L2L3-N'
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'GroundFault', 'on');
        case 'L1L2L3'
            set_param(faultBlockPath, 'FaultA', 'on');
            set_param(faultBlockPath, 'FaultB', 'on');
            set_param(faultBlockPath, 'FaultC', 'on');
            set_param(faultBlockPath, 'GroundFault', 'off');
        otherwise
            error('Invalid faulted phase: %s. Must be L1-N, L2-N, L3-N, L1L2-N...', scenario.faultedPhase);
    end
    
    % Specify scopes to log (assuming scopes are named 'VoltageScope' and 'CurrentScope')
    voltageScopeVF = [baseModel '/VoltageScopeVF'];
    currentScopeVF = [baseModel '/CurrentScopeVF'];
    voltageScopeHF = [baseModel '/VoltageScopeHF'];
    currentScopeHF = [baseModel '/CurrentScopeHF'];
    voltageScopeGF = [baseModel '/VoltageScopeGF'];
    currentScopeGF = [baseModel '/CurrentScopeGF'];
    voltageScopeGAF = [baseModel '/VoltageScopeGAF'];
    currentScopeGAF = [baseModel '/CurrentScopeGAF'];
    threePhaseToOne = [baseModel '/threePhaseToOne'];

    % Ensure signal logging is enabled
    set_param(baseModel, 'SignalLogging', 'on');
    set_param(baseModel, 'SignalLoggingName', 'logsout');
   
    % Ensure the simulation logs are configured
    set_param(baseModel, 'SignalLogging', 'on', 'SignalLoggingName', 'logsout');
    
    % Run the simulation
    try
        simOut = sim(baseModel, 'ReturnWorkspaceOutputs', 'on',"StopTime",stopTime);
        fprintf('Simulation completed for model: %s\n', baseModel);
    catch ME
        error('Simulation failed for model: %s. Error: %s', baseModel, ME.message);
    end
    
    % Extract the logged data from the To Workspace blocks
    try
        voltageStructVF = simOut.get('voltageVF');
        currentStructVF = simOut.get('currentVF');
        voltageStructHF = simOut.get('voltageHF');
        currentStructHF = simOut.get('currentHF');
        voltageStructGF = simOut.get('voltageGF');
        currentStructGF = simOut.get('currentGF');
        voltageStructGAF = simOut.get('voltageGAF');
        currentStructGAF = simOut.get('currentGAF'); 
        % Tmax = simOut.get('tmax'); 
      
       
        % Ensure data is in the form of a structure with Time and Data fields
        voltageDataVF = voltageStructVF.signals.values;
        currentDataVF = currentStructVF.signals.values;
        voltageDataHF = voltageStructHF.signals.values;
        currentDataHF = currentStructHF.signals.values;
        voltageDataGF = voltageStructGF.signals.values;
        currentDataGF = currentStructGF.signals.values;
        voltageDataGAF = voltageStructGAF.signals.values;
        currentDataGAF = currentStructGAF.signals.values;
        timeData = voltageStructVF.time; % Extract the time data
    catch ME
        error('Failed to extract data from simulation output in model: %s. Error: %s', baseModel, ME.message);
    end
    
    % Calculate the zero-sequence components
    zeroSeqVoltageVF = sum(voltageDataVF, 2)/3;
    zeroSeqCurrentVF = sum(currentDataVF, 2)/3;
    zeroSeqVoltageHF = sum(voltageDataHF, 2)/3;
    zeroSeqCurrentHF = sum(currentDataHF, 2)/3;
    zeroSeqVoltageGF = sum(voltageDataGF, 2)/3;
    zeroSeqCurrentGF = sum(currentDataGF, 2)/3;
    zeroSeqVoltageGAF = sum(voltageDataGAF, 2)/3;
    zeroSeqCurrentGAF = sum(currentDataGAF, 2)/3;

    %Calculate RMS Values
    zeroSeqVoltageRMSVF = rms(zeroSeqVoltageVF);
    zeroSeqCurrentRMSVF = rms(zeroSeqCurrentVF);
    zeroSeqVoltageRMSHF = rms(zeroSeqCurrentVF);
    zeroSeqCurrentRMSHF = rms(zeroSeqCurrentHF);
    zeroSeqVoltageRMSGF = rms(zeroSeqVoltageGF);
    zeroSeqCurrentRMSGF = rms(zeroSeqCurrentGF);
    zeroSeqVoltageRMSGAF = rms(zeroSeqVoltageGAF);
    zeroSeqCurrentRMSGAF = rms(zeroSeqCurrentGAF);

    %Maximale Amplitude des Signals
    maxzeroSeqVoltageVF = max(abs(zeroSeqVoltageVF));
    maxzeroSeqCurrentVF = max(abs(zeroSeqCurrentVF));
    maxzeroSeqVoltageHF = max(abs(zeroSeqVoltageHF));
    maxzeroSeqCurrentHF = max(abs(zeroSeqCurrentHF));
    maxzeroSeqVoltageGF = max(abs(zeroSeqVoltageGF));
    maxzeroSeqCurrentGF = max(abs(zeroSeqCurrentGF));
    maxzeroSeqVoltageGAF = max(abs(zeroSeqVoltageGAF));
    maxzeroSeqCurrentGAF = max(abs(zeroSeqCurrentGAF));

    %Winkel des Nullstromes und der Nullspannung
    angleTransformVoltageVF = hilbert(zeroSeqVoltageVF);
    angleTransformCurrentVF = hilbert(zeroSeqCurrentVF);
    angleTransformVoltageHF = hilbert(zeroSeqVoltageHF);
    angleTransformCurrentHF = hilbert(zeroSeqCurrentHF);
    angleTransformVoltageGF = hilbert(zeroSeqVoltageGF);
    angleTransformCurrentGF = hilbert(zeroSeqCurrentGF);
    angleTransformVoltageGAF = hilbert(zeroSeqVoltageGAF);
    angleTransformCurrentGAF = hilbert(zeroSeqCurrentGAF);

    angleTransformVoltageVF2 = zeroSeqVoltageVF + 1j*imag(angleTransformVoltageVF);
    angleTransformCurrentVF2 = zeroSeqCurrentVF + 1j*imag(angleTransformCurrentVF);
    angleTransformVoltageHF2 = zeroSeqVoltageHF + 1j*imag(angleTransformVoltageHF);
    angleTransformCurrentHF2 = zeroSeqCurrentHF + 1j*imag(angleTransformCurrentHF);
    angleTransformVoltageGF2 = zeroSeqVoltageGF + 1j*imag(angleTransformVoltageGF);
    angleTransformCurrentGF2 = zeroSeqCurrentGF + 1j*imag(angleTransformCurrentGF);
    angleTransformVoltageGAF2 = zeroSeqVoltageGAF + 1j*imag(angleTransformVoltageGAF);
    angleTransformCurrentGAF2 = zeroSeqCurrentGAF + 1j*imag(angleTransformCurrentGAF);

    angleZeroSeqVoltageVF = angle(angleTransformVoltageVF2);
    angleZeroSeqCurrentVF = angle(angleTransformCurrentVF2);
    angleZeroSeqVoltageHF = angle(angleTransformVoltageHF2);
    angleZeroSeqCurrentHF = angle(angleTransformCurrentHF2);
    angleZeroSeqVoltageGF = angle(angleTransformVoltageGF2);
    angleZeroSeqCurrentGF = angle(angleTransformCurrentGF2);
    angleZeroSeqVoltageGAF = angle(angleTransformVoltageGAF2);
    angleZeroSeqCurrentGAF = angle(angleTransformCurrentGAF2);
    
    %
    % firstNonNanValue=NaN;
    % for l=1:length(tMaxData)
    %     if ~isnan(tMaxData(l))
    %         firstNonNanValue=tMaxData(l);
    %         break;
    %     end
    % end
    % 

    %    % Number of phases
    % [numPhases, numSamples] = size(voltages);
    % 
    % % Initialize results structure
    % results = struct;
    % 
    % % Find zero-crossings (where the product of adjacent samples is negative)
    % zeroCrossMask = voltages(:, 1:end-1) .* voltages(:, 2:end) < 0;
    % 
    % % Find the first zero-crossing index for each phase
    % [~, firstZeroCrossIndex] = max(zeroCrossMask, [], 2);
    % 
    % % Initialize result arrays
    % firstZeroCrossTimes = nan(1, numPhases);
    % firstMaxTimes = nan(1, numPhases);
    % firstMaxValues = nan(1, numPhases);
    % 
    % % Vectorized handling for all phases
    % for phase = 1:numPhases
    %     if firstZeroCrossIndex(phase) > 0
    %         % Calculate the first zero-crossing time
    %         zeroCrossIdx = firstZeroCrossIndex(phase);
    %         firstZeroCrossTimes(phase) = time(zeroCrossIdx);
    % 
    %         % Extract the voltage segment starting after the first zero crossing
    %         segment = voltages(phase, zeroCrossIdx:end);
    % 
    %         % Find the first maximum in the segment
    %         [maxValue, maxIdx] = max(segment);
    % 
    %         % Store the maximum time and value
    %         firstMaxTimes(phase) = time(zeroCrossIdx + maxIdx - 1);
    %         firstMaxValues(phase) = maxValue;
    %     end
    % end
    % 
    % % Populate the results structure
    % for phase = 1:numPhases
    %     results.(sprintf('PhaseL%d', phase)) = struct(...
    %         'zeroCrossTime', firstZeroCrossTimes(phase), ...
    %         'maxTime', firstMaxTimes(phase), ...
    %         'maxValue', firstMaxValues(phase) ...
    %     );
    % end
    % 

   %Fehlerrichtungsalgorithmen
   %Vor Fehler 
   unenn=20000;%Nennspannung

  
% 
%    %cos phi Verfahren
%         %Einstellparameter
%         %Prozentuale Verlagerungsspannung Ansprechwert 
%         UNEP=0.3;
%         %Ansprechwert Wirkstrom in A
%         IEP=2;
%         %tot-bereich in °
%         phiTot=2;
%         %Verzögerung
%         tIEP=3000;
%         tIEP1=tIEP*5;
% 
%         %tot-Winkel
% 
%         upperDeadZone90= 90+phiTot;
%         lowerDeadZone90= 90-phiTot;
%         upperDeadZone270= 270+phiTot;
%         lowerDeadZone270= 270-phiTot;
% 
%         IEPBdirectionVF=0;
%         IEPAdirectionVF=0;
%         IEPnoDirectionVF=0;
%         IEPBdirectionHF=0;
%         IEPAdirectionHF=0;
%         IEPnoDirectionHF=0;
%         IEPBdirectionGF=0;
%         IEPAdirectionGF=0;
%         IEPnoDirectionGF=0;
%         IEPBdirectionGAF=0;
%         IEPAdirectionGAF=0;
%         IEPnoDirectionGAF=0;
% 
%          test10=abs(zeroSeqVoltageVF);
%         tschwelle1VF = find(abs(zeroSeqVoltageVF)>UNEP*unenn, 1);
%         tschwelle1HF = find(abs(zeroSeqVoltageHF)>UNEP*unenn, 1);
%         tschwelle1GF = find(abs(zeroSeqVoltageGF)>UNEP*unenn, 1);
%         tschwelle1GAF = find(abs(zeroSeqVoltageGAF)>UNEP*unenn, 1);
% 
%         if (tschwelle1VF>0)
%         %Winkel zwischen Summenstrom und Verlagerungsspannung vor Fehler
%         IEPangleUIVF=wrapTo360((rad2deg(angleZeroSeqVoltageVF-angleZeroSeqCurrentVF)));
%         IEPangleUIHF=wrapTo360((rad2deg(angleZeroSeqVoltageHF-angleZeroSeqCurrentHF)));
%         IEPangleUIGF=wrapTo360((rad2deg(angleZeroSeqVoltageGF-angleZeroSeqCurrentGF)));
%         IEPangleUIGAF=wrapTo360((rad2deg(angleZeroSeqVoltageGAF-angleZeroSeqCurrentGAF)));
% 
%         %Berechnung Wirkstrom
%         IwirkrestVF=3.*abs(zeroSeqCurrentVF).*cos(IEPangleUIVF);
%         IwirkrestHF=3.*abs(zeroSeqCurrentHF).*cos(IEPangleUIHF);
%         IwirkrestGF=3.*abs(zeroSeqCurrentGF).*cos(IEPangleUIGF);
%         IwirkrestGAF=3.*abs(zeroSeqCurrentGAF).*cos(IEPangleUIGAF);
% 
% 
%         %Ansprechverzögerung in ms
%         tAuswertungVF=tschwelle1VF+tIEP1;
%         tAuswertungHF=tschwelle1HF+tIEP1;
%         tAuswertungGF=tschwelle1GF+tIEP1;
%         tAuswertungGAF=tschwelle1GAF+tIEP1;
% 
%         %Auswertung zum definierten Zeitpunkt für Richtungsauswertung
%         IEPangleUIauswertungVF=IEPangleUIVF(tAuswertungVF);
%         IwirkrestAuswertungVF=IwirkrestVF(tAuswertungVF);
%         IEPangleUIauswertungHF=IEPangleUIHF(tAuswertungHF);
%         IwirkrestAuswertungHF=IwirkrestHF(tAuswertungHF);
%         IEPangleUIauswertungGF=IEPangleUIGF(tAuswertungGF);
%         IwirkrestAuswertungGF=IwirkrestGF(tAuswertungGF);
%         IEPangleUIauswertungGAF=IEPangleUIGAF(tAuswertungGAF);
%         IwirkrestAuswertungGAF=IwirkrestGAF(tAuswertungGAF);
% 
% 
% if IwirkrestAuswertungVF > abs(IEP)
%     % Check if the angle is within either dead zone
%     if (IEPangleUIauswertungVF > lowerDeadZone90 && IEPangleUIauswertungVF < upperDeadZone90) || ...
%        (IEPangleUIauswertungVF > lowerDeadZone270 && IEPangleUIauswertungVF < upperDeadZone270)
%         % If the angle is within the dead zones
%         IEPBdirectionVF = 0;
%         IEPAdirectionVF = 0;
%         IEPnoDirectionVF = 1;
%     else
%         % Check IwirkrestAuswertungVF conditions
%         if IwirkrestAuswertungVF < -IEP
%             % A direction
%             IEPBdirectionVF = 0;
%             IEPAdirectionVF = 1;
%             IEPnoDirectionVF = 0;
%         elseif IwirkrestAuswertungVF > IEP
%             % B direction
%             IEPBdirectionVF = 1;
%             IEPAdirectionVF = 0;
%             IEPnoDirectionVF = 0;
%         else
%             % No valid direction
%             IEPBdirectionVF = 0;
%             IEPAdirectionVF = 0;
%             IEPnoDirectionVF = 1;
%         end
%     end
% else
%     % If IwirkrestAuswertungVF is not greater than abs(IEP)
%     IEPBdirectionVF = 0;
%     IEPAdirectionVF = 0;
%     IEPnoDirectionVF = 1;
% end
% 
% if IwirkrestAuswertungHF > abs(IEP)
%     % Check if the angle is within either dead zone
%     if (IEPangleUIauswertungHF > lowerDeadZone90 && IEPangleUIauswertungHF < upperDeadZone90) || ...
%        (IEPangleUIauswertungHF > lowerDeadZone270 && IEPangleUIauswertungHF < upperDeadZone270)
%         % If the angle is within the dead zones
%         IEPBdirectionHF = 0;
%         IEPAdirectionHF = 0;
%         IEPnoDirectionHF = 1;
%     else
%         % Check IwirkrestAuswertungHF conditions
%         if IwirkrestAuswertungHF < -IEP
%             % A direction
%             IEPBdirectionHF = 0;
%             IEPAdirectionHF = 1;
%             IEPnoDirectionHF = 0;
%         elseif IwirkrestAuswertungHF > IEP
%             % B direction
%             IEPBdirectionHF = 1;
%             IEPAdirectionHF = 0;
%             IEPnoDirectionHF = 0;
%         else
%             % No valid direction
%             IEPBdirectionHF = 0;
%             IEPAdirectionHF = 0;
%             IEPnoDirectionHF = 1;
%         end
%     end
% else
%     % If IwirkrestAuswertungHF is not greater than abs(IEP)
%     IEPBdirectionHF = 0;
%     IEPAdirectionHF = 0;
%     IEPnoDirectionHF = 1;
% end
% 
%   if IwirkrestAuswertungGF > abs(IEP)
%     % Check if the angle is within either dead zone
%     if (IEPangleUIauswertungGF > lowerDeadZone90 && IEPangleUIauswertungGF < upperDeadZone90) || ...
%        (IEPangleUIauswertungGF > lowerDeadZone270 && IEPangleUIauswertungGF < upperDeadZone270)
%         % If the angle is within the dead zones
%         IEPBdirectionGF = 0;
%         IEPAdirectionGF = 0;
%         IEPnoDirectionGF = 1;
%     else
%         % Check IwirkrestAuswertungGF conditions
%         if IwirkrestAuswertungGF < -IEP
%             % A direction
%             IEPBdirectionGF = 0;
%             IEPAdirectionGF = 1;
%             IEPnoDirectionGF = 0;
%         elseif IwirkrestAuswertungGF > IEP
%             % B direction
%             IEPBdirectionGF = 1;
%             IEPAdirectionGF = 0;
%             IEPnoDirectionGF = 0;
%         else
%             % No valid direction
%             IEPBdirectionGF = 0;
%             IEPAdirectionGF = 0;
%             IEPnoDirectionGF = 1;
%         end
%     end
% else
%     % If IwirkrestAuswertungGF is not greater than abs(IEP)
%     IEPBdirectionGF = 0;
%     IEPAdirectionGF = 0;
%     IEPnoDirectionGF = 1;
% end
% 
% if IwirkrestAuswertungGAF > abs(IEP)
%     % Check if the angle is within either dead zone
%     if (IEPangleUIauswertungGAF > lowerDeadZone90 && IEPangleUIauswertungGAF < upperDeadZone90) || ...
%        (IEPangleUIauswertungGAF > lowerDeadZone270 && IEPangleUIauswertungGAF < upperDeadZone270)
%         % If the angle is within the dead zones
%         IEPBdirectionGAF = 0;
%         IEPAdirectionGAF = 0;
%         IEPnoDirectionGAF = 1;
%     else
%         % Check IwirkrestAuswertungGAF conditions
%         if IwirkrestAuswertungGAF < -IEP
%             % A direction
%             IEPBdirectionGAF = 0;
%             IEPAdirectionGAF = 1;
%             IEPnoDirectionGAF = 0;
%         elseif IwirkrestAuswertungGAF > IEP
%             % B direction
%             IEPBdirectionGAF = 1;
%             IEPAdirectionGAF = 0;
%             IEPnoDirectionGAF = 0;
%         else
%             % No valid direction
%             IEPBdirectionGAF = 0;
%             IEPAdirectionGAF = 0;
%             IEPnoDirectionGAF = 1;
%         end
%     end
% else
%     % If IwirkrestAuswertungGAF is not greater than abs(IEP)
%     IEPBdirectionGAF = 0;
%     IEPAdirectionGAF = 0;
%     IEPnoDirectionGAF = 1;
% end

%Wischerverfahren
%Einstellparameter
%Ansprechwert Verlagerungsspannung in Prozent
UNET = 30;
%Ansprechwert Erdschlusswischerstrom in A in RMS Werten
IET= 100;
%Ansprechwert ab dem 2.Erdwischerstrom in A
IETx=100;

%Anregeschwelle finden
UNETSchwelle=0.01*UNET*unenn;

% Hochpass Filter
Fs = 5000; % Sampling frequency in Hz (Adjust based on your data)
% Bandpass Filter Parameters
Fpass1 = 10; % Lower cutoff frequency in Hz
Fpass2 = 1000; % Upper cutoff frequency in Hz
order = 6; % Filter order

% Design the Bandpass Filter
bpFilter = designfilt('bandpassiir', ...
    'FilterOrder', order, ...
    'HalfPowerFrequency1', Fpass1, ...
    'HalfPowerFrequency2', Fpass2, ...
    'SampleRate', Fs);

% Visualize the Filter's Frequency Response
%fvtool(bpFilter); % Opens a tool to inspect the filter

% Apply the Filter
filteredVoltageWischerVF = filtfilt(bpFilter, zeroSeqVoltageVF); % Filter voltage zero sequence
filteredCurrentWischerVF = filtfilt(bpFilter, zeroSeqCurrentVF); % Filter current zero sequence


IETBdirectionVF=0;
IETAdirectionVF=0;
IETNoDirectionVF=0;
IETBdirectionHF=0;
IETAdirectionHF=0;
IETNoDirectionHF=0;
IETBdirectionGF=0;
IETAdirectionGF=0;
IETNoDirectionGF=0;
IETBdirectionGAF=0;
IETAdirectionGAF=0;
IETNoDirectionGAF=0;

tschwelleWischerVF = find(abs(filteredVoltageWischerVF)>UNETSchwelle, 1);
if(maxzeroSeqVoltageVF>UNETSchwelle & max(filteredCurrentWischerVF)>IET)
    wischerCurrentVF= filteredCurrentWischerVF(tschwelleWischerVF);
    wischerVoltageVF= filteredVoltageWischerVF(tschwelleWischerVF);
    polarityWischerCurrentVF=sign(wischerCurrentVF);
    polarityWischerVoltageVF=sign(wischerVoltageVF);
    if sign(polarityWischerVoltageVF) == sign(polarityWischerCurrentVF)
       IETBdirectionVF = 0;
       IETAdirectionVF = 1;
       IETNodirectionVF = 0;
    else
       IETBdirectionVF = 1;
       IETAdirectionVF = 0;
       IETNodirectionVF = 0;
    end
else
IETBdirectionVF = 0;
IETAdirectionVF = 0;
IETNoDirectionVF = 1;
end




% % Create the plot
% figure;
% yyaxis left;
% plot(timeData, filteredVoltageWischerVF, 'b', 'LineWidth', 1.5); % Voltage in blue
% ylabel('Zero-Sequence Voltage (pu)');
% 
% 
% yyaxis right;
% plot(timeData, filteredCurrentWischerVF, 'r', 'LineWidth', 1.5); % Current in red
% ylabel('Zero-Sequence Current (pu)');
% 
% % Set symmetric y-limits around zero
% yyaxis left;
% ylim([-1.3*max(filteredVoltageWischerVF), 1.3*max(filteredVoltageWischerVF)]);
% 
% yyaxis right;
% ylim([-1.3*max(filteredCurrentWischerVF), 1.3*max(filteredCurrentWischerVF)]);
% 
% % Customize plot
% xlabel('Time (s)');
% title('Zero-Sequence Voltage and Current');
% grid on;
% legend({'Zero-Sequence Voltage', 'Zero-Sequence Current'}, 'Location', 'best');


%Oberschwingungsverfahren
% Frequency of interest
targetFreq = 250; % 5th harmonic frequency (in Hz)
bandwidth = 10;   % Bandwidth around the target frequency (in Hz)

% Design a bandpass filter
[b, a] = butter(4, [(targetFreq - bandwidth/2) / (Fs/2), (targetFreq + bandwidth/2) / (Fs/2)], 'bandpass');

% Apply the bandpass filter to each zero-sequence signal
filteredVoltageVF = filtfilt(b, a, zeroSeqVoltageVF);
filteredCurrentVF = filtfilt(b, a, zeroSeqCurrentVF);

filteredVoltageHF = filtfilt(b, a, zeroSeqVoltageHF);
filteredCurrentHF = filtfilt(b, a, zeroSeqCurrentHF);

filteredVoltageGF = filtfilt(b, a, zeroSeqVoltageGF);
filteredCurrentGF = filtfilt(b, a, zeroSeqCurrentGF);

filteredVoltageGAF = filtfilt(b, a, zeroSeqVoltageGAF);
filteredCurrentGAF = filtfilt(b, a, zeroSeqCurrentGAF);
% Filtered signals: Voltage and Current (Example for VF)
voltageVF = filteredVoltageVF;
currentVF = filteredCurrentVF;
voltageHF = filteredVoltageHF;
currentHF = filteredCurrentHF;

% Calculate instantaneous phases using the Hilbert transform
voltagePhaseVF = angle(hilbert(voltageVF));
currentPhaseVF = angle(hilbert(currentVF));
voltagePhaseHF = angle(hilbert(voltageHF));
currentPhaseHF = angle(hilbert(currentHF));

% Delta angle (in radians)
deltaAngleVF = voltagePhaseVF - currentPhaseVF;
deltaAngleHF = voltagePhaseHF - currentPhaseHF;

% Convert delta angle to degrees
deltaAngleDegreesVF = rad2deg(deltaAngleVF);
deltaAngleDegreesHF = rad2deg(deltaAngleHF);

% Wrap angles to the range [-180, 180]
deltaAngleDegreesVF = mod(deltaAngleDegreesVF + 180, 360) - 180;
deltaAngleDegreesHF = mod(deltaAngleDegreesHF + 180, 360) - 180;

% Plot the delta angle
figure;
plot(deltaAngleDegreesVF);
title('Delta Angle (Filtered Zero-Sequence Voltage and Current, VF)');
xlabel('Sample');
ylabel('Delta Angle (degrees)');
figure;
plot(deltaAngleDegreesHF);
title('Delta Angle (Filtered Zero-Sequence Voltage and Current, HF)');
xlabel('Sample');
ylabel('Delta Angle (degrees)');


    % Store the data in the structure
    scenarioData = struct();
    scenarioData.VoltageDataVF = voltageDataVF;
    scenarioData.CurrentDataVF = currentDataVF;
    scenarioData.VoltageDataHF = voltageDataHF;
    scenarioData.CurrentDataHF = currentDataHF;
    scenarioData.VoltageDataGF = voltageDataGF;
    scenarioData.CurrentDataGF = currentDataGF;
    scenarioData.VoltageDataGAF = voltageDataGAF;
    scenarioData.CurrentDataGAF = currentDataGAF;
    scenarioData.ZeroSeqVoltageVF = zeroSeqVoltageVF;
    scenarioData.ZeroSeqCurrentVF = zeroSeqCurrentVF;
    scenarioData.ZeroSeqVoltageHF = zeroSeqVoltageHF;
    scenarioData.ZeroSeqCurrentHF = zeroSeqCurrentHF;
    scenarioData.ZeroSeqVoltageGF = zeroSeqVoltageGF;
    scenarioData.ZeroSeqCurrentGF = zeroSeqCurrentGF;
    scenarioData.ZeroSeqVoltageGAF = zeroSeqVoltageGAF;
    scenarioData.ZeroSeqCurrentGAF = zeroSeqCurrentGAF;
    scenarioData.maxZeroSeqVoltageVF = maxzeroSeqVoltageVF;
    scenarioData.maxZeroSeqCurrentVF = maxzeroSeqCurrentVF;
    scenarioData.maxZeroSeqVoltageHF = maxzeroSeqVoltageHF;
    scenarioData.maxZeroSeqCurrentHF = maxzeroSeqCurrentHF;
    scenarioData.maxZeroSeqVoltageGF = maxzeroSeqVoltageGF;
    scenarioData.maxZeroSeqCurrentGF = maxzeroSeqCurrentGF;
    scenarioData.maxZeroSeqVoltageGAF = maxzeroSeqVoltageGAF;
    scenarioData.maxZeroSeqCurrentGAF = maxzeroSeqCurrentGAF;
    scenarioData.TimeData = timeData;
    % scenarioData.TMaxData = tMaxData;
    % scenarioData.FirstNonNanValue = firstNonNanValue;
    
    %Storing Direction Data
    % scenarioData.IEPAngleUIauswertungVF = IEPangleUIauswertungVF;
    % scenarioData.IEPAngleUIauswertungHF = IEPangleUIauswertungHF;
    % scenarioData.IEPAngleUIauswertungGF = IEPangleUIauswertungGF;
    % scenarioData.IEPAngleUIauswertungGAF = IEPangleUIauswertungGAF;
    % scenarioData.IWirkrestAuswertungVF = IwirkrestAuswertungVF;
    % scenarioData.IWirkrestAuswertungHF = IwirkrestAuswertungHF;
    % scenarioData.IWirkrestAuswertungGF = IwirkrestAuswertungGF;
    % scenarioData.IWirkrestAuswertungGAF = IwirkrestAuswertungGAF;
    % scenarioData.IEPBDirectionVF = IEPBdirectionVF;
    % scenarioData.IEPADirectionVF = IEPAdirectionVF;
    % scenarioData.IEPNoDirectionVF = IEPnoDirectionVF;
    % scenarioData.IEPBDirectionHF = IEPBdirectionHF;
    % scenarioData.IEPADirectionHF = IEPAdirectionHF;
    % scenarioData.IEPNoDirectionHF = IEPnoDirectionHF;
    % scenarioData.IEPBDirectionGF = IEPBdirectionGF;
    % scenarioData.IEPADirectionGF = IEPAdirectionGF;
    % scenarioData.IEPNoDirectionGF = IEPnoDirectionGF;
    % scenarioData.IEPBDirectionGAF = IEPBdirectionGAF;
    % scenarioData.IEPADirectionGAF = IEPAdirectionGAF;
    % scenarioData.IEPNoDirectionGAF = IEPnoDirectionGAF;
    scenarioData.IETBDirectionVF = IETBdirectionVF;
    scenarioData.IETADirectionVF = IETAdirectionVF;
    scenarioData.IETNoDirectionVF = IETNoDirectionVF;
   
    % Define parameters for plotting
     Fs = 1 / (timeData(2) - timeData(1));  % Sampling frequency
     T = 1 / Fs;                            % Sampling period
     L = numel(zeroSeqVoltageVF);                % Length of signal
     f = Fs*(0:(L/2))/L;                    % Frequency range for plotting
    
    % Maximale Amplituden
    maxZeroSeqVoltageVF(i) = max(abs(zeroSeqVoltageVF));
    maxZeroSeqCurrentVF(i) = max(abs(zeroSeqCurrentVF));
    maxZeroSeqVoltageHF(i) = max(abs(zeroSeqVoltageHF));
    maxZeroSeqCurrentHF(i) = max(abs(zeroSeqCurrentHF));
    maxZeroSeqVoltageGF(i) = max(abs(zeroSeqVoltageGF));
    maxZeroSeqCurrentGF(i) = max(abs(zeroSeqCurrentGF));
    maxZeroSeqVoltageGAF(i) = max(abs(zeroSeqVoltageGAF));
    maxZeroSeqCurrentGAF(i) = max(abs(zeroSeqCurrentGAF));


    % Add the scenario data to the main structure 
     allScenarioData.(sprintf('Scenario_%d', i)) = scenarioData;
   
    % Close the modified model without saving changes
    try
        close_system(newModel, 0);
        close_system (baseModel, 0);
        fprintf('Model %s closed successfully.\n', newModel);
    catch ME
        error('Failed to close model: %s. Error: %s', newModel, ME.message);
    end
    %FFT
end
  

%Auswertung Fehlerrichtungsanzeige / Wie viele Richtungsereignisse
    %wurden korrekt ausgewertet?
    
% %cos phi
% % Messpunkt vor dem Fehler
% %B Richtung vor Fehler korrekt ausgewertet
% IEPBdirectionVFAuswertung=sum([scenarioData.IEPBDirectionVF]);
% % A Richtung vor Fehler falsch ausgewertet
% IEPAdirectionVFAuswertung=sum([scenarioData.IEPADirectionVF]);
% % Keine Richtung vor Fehler ausgewertet
% IEPnoDirectionVFAuswertung=sum([scenarioData.IEPNoDirectionVF]);
% 
% % Messpunkt hinter dem Fehler
% %B Richtung vor Fehler falsch ausgewertet
% IEPBdirectionHFAuswertung=sum([scenarioData.IEPBDirectionHF]);
% % A Richtung vor Fehler richtig ausgewertet
% IEPAdirectionHFAuswertung=sum([scenarioData.IEPADirectionHF]);
% % Keine Richtung vor Fehler ausgewertet
% IEPnoDirectionHFAuswertung=sum([scenarioData.IEPNoDirectionHF]);
% 
% % Messpunkt Generatornah
% %B Richtung vor Fehler korrekt ausgewertet
% IEPBdirectionGFAuswertung=sum([scenarioData.IEPBDirectionGF]);
% % A Richtung vor Fehler falsch ausgewertet
% IEPAdirectionGFAuswertung=sum([scenarioData.IEPADirectionGF]);
% % Keine Richtung vor Fehler ausgewertet
% IEPnoDirectionGFAuswertung=sum([scenarioData.IEPNoDirectionGF]);
% 
% % Messpunkt hinter dem Fehler
% %B Richtung vor Fehler falsch ausgewertet
% IEPBdirectionGAFAuswertung=sum([scenarioData.IEPBDirectionGAF]);
% % A Richtung vor Fehler richtig ausgewertet
% IEPAdirectionGAFAuswertung=sum([scenarioData.IEPADirectionGAF]);
% % Keine Richtung vor Fehler ausgewertet
% IEPnoDirectionGAFAuswertung=sum([scenarioData.IEPNoDirectionGAF]);
% 
% % Calculate percentages for B and A directions
% if IEPBdirectionVFAuswertung > 0
%     IEPBdirectionVFAuswertung_Percentage = (IEPBdirectionVFAuswertung / numScenarios) * 100;
% else
%     IEPBdirectionVFAuswertung_Percentage = 0;
% end
% 
% if IEPAdirectionVFAuswertung > 0
%     IEPAdirectionVFAuswertung_Percentage = (IEPAdirectionVFAuswertung / numScenarios) * 100;
% else
%     IEPAdirectionVFAuswertung_Percentage = 0;
% end
% 
% if IEPnoDirectionVFAuswertung > 0
%     IEPnoDirectionVFAuswertung_Percentage = (IEPnoDirectionVFAuswertung / numScenarios) * 100;
% else
%     IEPnoDirectionVFAuswertung_Percentage = 0;
% end
% 
% % Display the results in a pop-up window
% IEPVFB_Direction_Result = sprintf('B direction correctly detected in %d/%d cases (%.2f%%).', IEPBdirectionGFAuswertung, numScenarios, IEPBdirectionVFAuswertung_Percentage);
% IEPVFA_Direction_Result = sprintf('A direction correctly detected in %d/%d cases (%.2f%%).', IEPnoDirectionGFAuswertung, numScenarios, IEPAdirectionVFAuswertung_Percentage);
% 
% msgbox({IEPVFB_Direction_Result, IEPVFA_Direction_Result}, 'Direction Detection Results');
