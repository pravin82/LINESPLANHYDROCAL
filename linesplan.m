filename='MOULDED OFFSET.xlsx';
data = dataset('xlsfile',filename);
data=dataset2cell(data);
%stations
stn=[0 0.25 0.5 1 1.5 2 3 4 7 8 8.5 9 9.5 9.75 10 ];
col='B':'S';
for i=1:length(col)
    Lloc(i)=xlsread(filename,'sheet1',strcat(col(i),'19'));
end
%WL0
for i=2:16
    WL0(i-1)=data{i,2};
end
%WL0.5
for i=2:16
    WL0_5(i-1)=data{i,3};
end
%WL1
for i=2:16
    WL1(i-1)=data{i,4};
end
%WL2
for i=2:16
    WL2(i-1)=data{i,5};
end
%WL3
for i=2:16
    WL3(i-1)=data{i,6};
end
%WL4
for i=2:16
    WL4(i-1)=data{i,7};
end
%WL5
for i=2:16
    WL5(i-1)=data{i,8};
end
%WL6
for i=2:16
    WL6(i-1)=data{i,9};
end
%WL7
for i=2:16
    WL7(i-1)=data{i,10};
end
%WL8
for i=2:16
    WL8(i-1)=data{i,11};
end
%WL9
for i=2:16
    WL9(i-1)=data{i,12};
end
%WL10
for i=2:16
    WL10(i-1)=data{i,13};
end
%WL10_45
for i=2:16
    WL10_45(i-1)=data{i,14};
end
%WL12
for i=2:16
    WL12(i-1)=data{i,15};
end
%WL14
for i=2:16
    WL14(i-1)=data{i,16};
end
%WL16
for i=2:16
    WL16(i-1)=data{i,17};
end
%WL18
for i=2:16
    WL18(i-1)=data{i,18};
end
%WL20
for i=2:16
    WL20(i-1)=data{i,19};
end


%HYDROSTATIC CALCULATIONS
%common columns
SM=[.25 1 .5 1 .75 2 1 2 1 2 .75 1 .5 1 .25];
LLMS=[-5 -4.5 -4 -3.5 -3 -2 -1 0 1 2 3 3.5 4 4.5 5];
h=17.4;
%section area


%WL0
Z=0;%change
filename='WL0';%change
secarea=SM.*WL0;%change
secarea=secarea*(Z/2);
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL0;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(1)*bdthwl(9)*2*Z);%change
%amid
b0=bdthwl(9);
amid=0;
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(1)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);%check
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%WL0_5
Z=0.5;%change
filename='WL0_5';%change
secarea=SM.*WL0_5;%change
secarea=secarea*(Z/2);
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL0_5;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(2)*bdthwl(9)*2*Z);%change
%amid
b1=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
cw=awp/(Lloc(2)*bdthwl(9)*2); %change
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(184*22.9);
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


%WL1
Z=1;%change
filename='WL1';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL1;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(3)*bdthwl(9)*Z*2);%change
%amid
b2=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(3)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL2
Z=2;%change
filename='WL2';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL2;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(4)*bdthwl(9)*Z*2);%change
%amid
%amid
b3=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(4)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL3
Z=3;%change
filename='WL3';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL3;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(5)*bdthwl(9)*Z*2);%change
%amid
b4=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(5)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL4
Z=4;%change
filename='WL4';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL4;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(6)*bdthwl(9)*Z*2);%change
%amid
%amid
b5=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(6)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL5
Z=5;%change
filename='WL5';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL5;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(7)*bdthwl(9)*Z*2);%change
%amid
%amid
b6=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(7)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL6
Z=6;%change
filename='WL6';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL6;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(8)*bdthwl(9)*Z*2);%change
%amid
%amid
b7=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(8)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL7
Z=7;%change
filename='WL7';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL7;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(9)*bdthwl(9)*Z*2);%change
%amid
%amid
b8=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(9)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL8
Z=8;%change
filename='WL8';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL8;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(10)*bdthwl(9)*Z*2);%change
%amid
%amid
b9=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(10)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL9
Z=9;%change
filename='WL9';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL9;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(11)*bdthwl(9)*Z*2);%change
%amid
%amid
b10=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(11)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL10
Z=10;%change
filename='WL10';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL10;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(12)*bdthwl(9)*Z*2);%change
%amid
%amid
b11=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(12)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL10_45
Z=10.45;%change
filename='WL10_45';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL10_45;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(13)*bdthwl(9)*Z*2);%change
%amid
%amid
b12=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(13)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL12
Z=12;%change
filename='WL12';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
secarea=((WL10+WL12)*(0.5)*2)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL12;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(14)*bdthwl(9)*Z*2);%change
%amid
%amid
b13=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
amid=amid+((b12+b13)*0.5*2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(14)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%WL14
Z=14;%change
filename='WL14';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
secarea=((WL10+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL12)*(0.5)*2)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL14;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(15)*bdthwl(9)*Z*2);%change
%amid
%amid
b14=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
amid=amid+((b12+b13)*0.5*2);
amid=amid+((b14+b13)*0.5*2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(15)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL16
Z=16;%change
filename='WL16';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
secarea=((WL10+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL16)*(0.5)*2)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL16;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(16)*bdthwl(9)*Z*2);%change
%amid
%amid
b15=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
amid=amid+((b12+b13)*0.5*2);
amid=amid+((b14+b13)*0.5*2);
amid=amid+((b14+b15)*0.5*2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(16)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


%WL18
Z=18;%change
filename='WL18';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
secarea=((WL10+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL16)*(0.5)*2)+secarea;
secarea=((WL18+WL16)*(0.5)*2)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL18;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(17)*bdthwl(9)*Z*2);%change
%amid
%amid
b16=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
amid=amid+((b12+b13)*0.5*2);
amid=amid+((b14+b13)*0.5*2);
amid=amid+((b14+b15)*0.5*2);
amid=amid+((b16+b15)*0.5*2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(17)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%WL20
Z=20;%change
filename='WL20';%change
%secarea
seacrea=WL0_5*(0.5/2);
secarea=((WL0_5+WL1)*(0.5/2))+secarea;
secarea=((WL2+WL1)*(0.5))+secarea;
secarea=((WL2+WL3)*(0.5))+secarea;
secarea=((WL4+WL3)*(0.5))+secarea;
secarea=((WL4+WL5)*(0.5))+secarea;
secarea=((WL6+WL5)*(0.5))+secarea;
secarea=((WL6+WL7)*(0.5))+secarea;
secarea=((WL8+WL7)*(0.5))+secarea;
secarea=((WL8+WL9)*(0.5))+secarea;
secarea=((WL10+WL9)*(0.5))+secarea;
secarea=((WL10+WL10_45)*(0.5)*.45)+secarea;
secarea=((WL10+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL12)*(0.5)*2)+secarea;
secarea=((WL14+WL16)*(0.5)*2)+secarea;
secarea=((WL18+WL16)*(0.5)*2)+secarea;
secarea=((WL18+WL20)*(0.5)*2)+secarea;
fnvol=SM.*secarea;
fnlongmom=fnvol.*LLMS;
secareamom=secarea.*LLMS;
fnofvermom=SM.*secareamom;
bdthwl=WL20;%change
fnofwparea=bdthwl.*SM;
fnofwpmom=fnofwparea.*LLMS;
fnlongMI=fnofwpmom.*LLMS;
bdth3=bdthwl.*bdthwl.*bdthwl;
fntransMI=SM.*bdth3;
tabi=[stn' SM' secarea' fnvol' LLMS' fnlongmom' secareamom' fnofvermom' bdthwl' fnofwparea' fnofwpmom' fnlongMI' bdth3' fntransMI'];
tabi=num2cell(tabi);
del=0;
for i=1:15
    if isnan(fnvol(i))==0
     del=del+4/3*h*fnvol(i);
    end
end
delsw=del*1.025;
delext=del*1.033;

cb=del/(Lloc(18)*bdthwl(9)*Z*2);%change
%amid
b17=bdthwl(9);%change
amid=0;
amid=amid+((b1+b0)*0.5/2);
amid=amid+((b1+b2)*0.5/2);
amid=amid+((b3+b2)*0.5);
amid=amid+((b3+b4)*0.5);
amid=amid+((b4+b5)*0.5);
amid=amid+((b6+b5)*0.5);
amid=amid+((b6+b7)*0.5);
amid=amid+((b8+b7)*0.5);
amid=amid+((b8+b9)*0.5);
amid=amid+((b10+b9)*0.5);
amid=amid+((b10+b11)*0.5);
amid=amid+((b12+b11)*0.5*.45);
amid=amid+((b12+b13)*0.5*2);
amid=amid+((b14+b13)*0.5*2);
amid=amid+((b14+b15)*0.5*2);
amid=amid+((b16+b15)*0.5*2);
amid=amid+((b16+b17)*0.5*2);
cm=amid/(bdthwl(9)*Z);
cp=cb/cm;
lmom=0;
for i=1:15
    if isnan(fnlongmom(i))==0
     lmom=lmom+4/3*h^2*fnlongmom(i);
    end
end
lcb=lmom/del;
vmom=0;
for i=1:15
    if isnan(fnofvermom(i))==0
     vmom=vmom+4/3*h^2*fnofvermom(i);
    end
end
vcb=vmom/del;
awp=0;
for i=1:15
    if isnan(fnofwparea(i))==0
     awp=awp+4/3*h*fnofwparea(i);
    end
end
tcp=1.025*awp/100;
cw=awp/(Lloc(18)*bdthwl(9)*2); %change
wpmom=0;
for i=1:15
    if isnan(fnofwpmom(i))==0
     wpmom=wpmom+4/3*h^2*fnofwpmom(i);
    end
end
lcf=wpmom/awp;
x=awp*lcf^2;
il=0;
for i=1:15
    if isnan(fnlongMI(i))==0
     il=il+4/3*h^3*fnlongMI(i);
    end
end
il=il-x;
bml=il/del;
it=0;
for i=1:15
    if isnan(fntransMI(i))==0
     it=it+4/3*h*fntransMI(i);
    end
end
bmt=it/del;
mct=(1.025*il)/(100*184);
tab2={del,delsw,delext,cb,cm,cp,lmom,lcb,vmom,vcb,awp,tcp,cw,wpmom,lcf,il,bml,it,bmt,mct};
xlswrite(filename,tab2,'sheet','A17')

header={'STATION','S.M','1/2SEC AREA','FN OF VOLUME','LONG LEVER FROM MIDSHIP','FN OF LONG. MOMENT','1/2 SEC AREA MOMENT','FN OF VERTICAL MOMENT','1/2 BREADTH OF W.L','FN OF W.P. AREA','FN OF W.P. MOMOMENT','FN OF LONG. MI','1/2 BREADTH^3','FN OF TRANS MI'};
xlswrite(filename,header,'sheet');
xlswrite(filename,tabi,'sheet','A2');
header2={'vol','weightS.W','weightEXT','CB','CM','CP','L-MOM','LCB','V-MOM','VCB','AWP','TPCm','CW','W.P.MOM','LCF','IL','BML','IT','BMT','MCT1cm'};
xlswrite(filename,header2,'sheet','A16')
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%







