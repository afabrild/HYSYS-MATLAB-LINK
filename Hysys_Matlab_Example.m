%% MATLAB - HYSYS LINK EXAMPLE
 % This code provides an example of direct link between ASPEN HYSYS and
 % MATLAB
 
 %Author: ANDRÉS FELIPE ABRIL DURAN
 %Afiliation: NATIONAL UNIVERSITY OF COLOMBIA
 %            CHEMICAL AND ENVIRONMENTAL ENGINEERING DEPARTMENT
 %e-mail: afabrild@unal.edu.co; afabrild@gmail.com
 
% MATLAB-HYSYS CONNECTION

Hysys=actxserver('Hysys.Application.V10.0'); %COM Technology
[stat,mess]=fileattrib; %getting actual folder name
simcase = Hysys.SimulationCases.Open([mess.Name '\Distill_Example.hsc']); %Opening HYSYS Simulation (Distill_Example.hsc)
simcase.invoke('Activate'); %Opening HYSYS Simulation (Distill_Example.hsc)

%
fs=simcase.get('flowsheet'); %accessing to flowsheet
op=fs.get('Operations'); %accessing to simulated operations (e.g Distillation column)
ms=fs.get('MaterialStreams'); %accessing to simulated material streams
%ns=fs.get('EnergyStreams'); % Energy streams
sheet = op.Item('Sheet_1'); %Spreadsheets (e.g "Sheet_1")
HySolver = simcase.Solver; %Hysys can solve? De/Activating the solver

%%Sensibility analysis example

Number_Trays = 90:5:150;

for i = 10:1:length(Number_Trays)
tic;
HySolver.CanSolve = 0;
HySolver.CanSolve = 1;
op.Item('C-101').ColumnFlowsheet.Operations.Item("Main_TS").NumberOfTrays = Number_Trays(i);
Feed_Location = op.Item('C-101').ColumnFlowsheet.FeedStreams.Item('Feed4');
op.Item('C-101').ColumnFlowsheet.Operations.Item("Main_TS").SpecifyFeedLocation(Feed_Location, 13);
Side_Location = op.Item('C-101').ColumnFlowsheet.LiquidProducts.Item('116');
op.Item('C-101').ColumnFlowsheet.Operations.Item("Main_TS").SpecifyDrawLocation(Side_Location, 19);
HySolver.CanSolve = 0;
HySolver.CanSolve = 1;
op.Item('C-101').ColumnFlowsheet.Run();
Column_Converge = op.Item('C-101').ColumnFlowsheet.CfsConverged;
Duty = sheet.Cell('B2').CellValue;
Steam_Cost = sheet.Cell('I4').CellValue*8000;
Cooling_Water_Cost = sheet.Cell('J4').CellValue*8000;
%saving information
Convergency (i) = Column_Converge;
Steam_Cost_Vector(i) = Steam_Cost;
Cooling_Cost_Vector(i) = Cooling_Water_Cost;
Reflux_Vector(i) = op.Item('C-101').ColumnFlowsheet.RefluxRatio;
Steam_Cons_Vector(i) = Duty;
toc;
end

%Plottin results
figure(1)
plot(Number_Trays, Steam_Cost_Vector, 'ok');
figure(2)
plot(Number_Trays, Cooling_Cost_Vector, 'ok');
figure(3)
plot(Number_Trays, Reflux_Vector, 'ok');
figure(4)
plot(Number_Trays, Steam_Cons_Vector, 'ok');




