clear all;
tic;

%%处理MRIO数据
%读整体
[Type Sheet Format]=xlsfinfo('WEC nexus data_2012.xlsx');
biao = zeros(1302,1302);
biao = xlsread('WEC nexus data_2012.xlsx','Output_Raw','C4:AXD1305');
biao_Y = zeros(1302,155);
biao_Y = xlsread('WEC nexus data_2012.xlsx','Output_Raw','AXE4:BDC1305');

%读取31层全部数据
for i = 1:31
    RIO(:,:,i)= biao(:,1+42*(i-1):42*i);
end

%整合列
RIO_tmp = zeros(1302,9,31);
for i = 1:31
    RIO_tmp(:,1,i)= RIO(:,1,i);
    RIO_tmp(:,2,i)= RIO(:,2,i)+RIO(:,3,i)+RIO(:,4,i)+RIO(:,5,i);
    RIO_tmp(:,3,i)= RIO(:,6,i)+RIO(:,7,i)+RIO(:,8,i)+RIO(:,9,i)+RIO(:,10,i)+RIO(:,11,i)+RIO(:,12,i);
    RIO_tmp(:,4,i)= RIO(:,13,i)+RIO(:,14,i)+RIO(:,15,i)+RIO(:,16,i)+RIO(:,17,i)+RIO(:,18,i)+RIO(:,19,i)+RIO(:,20,i)+RIO(:,21,i)+RIO(:,22,i)+RIO(:,23,i)+RIO(:,24,i);
    RIO_tmp(:,5,i)= RIO(:,25,i)+RIO(:,26,i)+RIO(:,27,i);
    RIO_tmp(:,6,i)= RIO(:,28,i);
    RIO_tmp(:,8,i)= RIO(:,31,i)+RIO(:,29,i);
    RIO_tmp(:,7,i)= RIO(:,32,i)+RIO(:,30,i);
    RIO_tmp(:,9,i)= RIO(:,33,i)+RIO(:,34,i)+RIO(:,35,i)+RIO(:,36,i)+RIO(:,37,i)+RIO(:,38,i)+RIO(:,39,i)+RIO(:,40,i)+RIO(:,41,i)+RIO(:,42,i);
end

%还原处理列之后的表
biao_lie = zeros(1302,9*31);
for i = 1:31
   biao_lie(:,1+9*(i-1):9*i) = RIO_tmp(:,:,i);
end

%读取biao_lie数据处理行
RIO_new = zeros(42,279,31);
for i = 1:31
    RIO_new(:,:,i)= biao_lie(1+42*(i-1):42*i,:);
end

%整合行
RIO_new_tmp = zeros(9,279,31);
for i = 1:31
    RIO_new_tmp(1,:,i)= RIO_new(1,:,i);
    RIO_new_tmp(2,:,i)= RIO_new(2,:,i)+RIO_new(3,:,i)+RIO_new(4,:,i)+RIO_new(5,:,i);
    RIO_new_tmp(3,:,i)= RIO_new(6,:,i)+RIO_new(7,:,i)+RIO_new(8,:,i)+RIO_new(9,:,i)+RIO_new(10,:,i)+RIO_new(11,:,i)+RIO_new(12,:,i);
    RIO_new_tmp(4,:,i)= RIO_new(13,:,i)+RIO_new(14,:,i)+RIO_new(15,:,i)+RIO_new(16,:,i)+RIO_new(17,:,i)+RIO_new(18,:,i)+RIO_new(19,:,i)+RIO_new(20,:,i)+RIO_new(21,:,i)+RIO_new(22,:,i)+RIO_new(23,:,i)+RIO_new(24,:,i);
    RIO_new_tmp(5,:,i)= RIO_new(25,:,i)+RIO_new(26,:,i)+RIO_new(27,:,i);
    RIO_new_tmp(6,:,i)= RIO_new(28,:,i);
    RIO_new_tmp(8,:,i)= RIO_new(31,:,i)+RIO_new(29,:,i);
    RIO_new_tmp(7,:,i)= RIO_new(32,:,i)+RIO_new(30,:,i);
    RIO_new_tmp(9,:,i)= RIO_new(33,:,i)+RIO_new(34,:,i)+RIO_new(35,:,i)+RIO_new(36,:,i)+RIO_new(37,:,i)+RIO_new(38,:,i)+RIO_new(39,:,i)+RIO_new(40,:,i)+RIO_new(41,:,i)+RIO_new(42,:,i);
end

biao_Y_new = zeros(279,155);
for i = 1:31
    biao_Y_new(1+9*(i-1),:)= biao_Y(1+42*(i-1),:);
    biao_Y_new(2+9*(i-1),:)= biao_Y(2+42*(i-1),:)+biao_Y(3+42*(i-1),:)+biao_Y(4+42*(i-1),:)+biao_Y(5+42*(i-1),:);
    biao_Y_new(3+9*(i-1),:)= biao_Y(6+42*(i-1),:)+biao_Y(7+42*(i-1),:)+biao_Y(8+42*(i-1),:)+biao_Y(9+42*(i-1),:)+biao_Y(10+42*(i-1),:)+biao_Y(11+42*(i-1),:)+biao_Y(12+42*(i-1),:);
    biao_Y_new(4+9*(i-1),:)= biao_Y(13+42*(i-1),:)+biao_Y(14+42*(i-1),:)+biao_Y(15+42*(i-1),:)+biao_Y(16+42*(i-1),:)+biao_Y(17+42*(i-1),:)+biao_Y(18+42*(i-1),:)+biao_Y(19+42*(i-1),:)+biao_Y(20+42*(i-1),:)+biao_Y(21+42*(i-1),:)+biao_Y(22+42*(i-1),:)+biao_Y(23+42*(i-1),:)+biao_Y(24+42*(i-1),:);
    biao_Y_new(5+9*(i-1),:)= biao_Y(25+42*(i-1),:)+biao_Y(26+42*(i-1),:)+biao_Y(27+42*(i-1),:);
    biao_Y_new(6+9*(i-1),:)= biao_Y(28+42*(i-1),:);
    biao_Y_new(8+9*(i-1),:)= biao_Y(31+42*(i-1),:)+biao_Y(29+42*(i-1),:);
    biao_Y_new(7+9*(i-1),:)= biao_Y(32+42*(i-1),:)+biao_Y(30+42*(i-1),:);
    biao_Y_new(9+9*(i-1),:)= biao_Y(33+42*(i-1),:)+biao_Y(34+42*(i-1),:)+biao_Y(35+42*(i-1),:)+biao_Y(36+42*(i-1),:)+biao_Y(37+42*(i-1),:)+biao_Y(38+42*(i-1),:)+biao_Y(39+42*(i-1),:)+biao_Y(40+42*(i-1),:)+biao_Y(41,:)+biao_Y(42+42*(i-1),:);
end

%还原处理行之后的表
biao_hang = zeros(279,279);
for i = 1:31
   biao_hang(1+9*(i-1):9*i,:) = RIO_new_tmp(:,:,i);
end

biao_hang_carbon_energy(1:9*25,1:9*25) = biao_hang(1:9*25,1:9*25);
biao_hang_carbon_energy(1:9*25,9*25+1:30*9) = biao_hang(1:9*25,9*26+1:31*9);
biao_hang_carbon_energy(9*25+1:30*9,1:9*25) = biao_hang(9*26+1:31*9,1:9*25);
biao_hang_carbon_energy(9*25+1:30*9,9*25+1:30*9) = biao_hang(9*26+1:31*9,9*26+1:31*9);
biao_Y_carbon_energy(1:9*25,1:5*25) = biao_Y_new(1:9*25,1:5*25);
biao_Y_carbon_energy(1:9*25,5*25+1:30*5) = biao_Y_new(1:9*25,5*26+1:31*5);
biao_Y_carbon_energy(9*25+1:30*9,1:5*25) = biao_Y_new(9*26+1:31*9,1:5*25);
biao_Y_carbon_energy(9*25+1:30*9,5*25+1:30*5) = biao_Y_new(9*26+1:31*9,5*26+1:31*5);

%读取并处理总产出X及最终使用Y
ZCC = zeros(1302,1);
ZZSY = zeros(1302,1);
ZCC = xlsread('WEC nexus data_2012.xlsx','Output_Raw','BDE4:BDE1305');
ZZSY = xlsread('WEC nexus data_2012.xlsx','Output_Raw','BDD4:BDD1305');

X = zeros(279,1);
Y = zeros(279,1);
for i = 1:31
    X(1+9*(i-1),:)= ZCC(1+42*(i-1),:);
    X(2+9*(i-1),:)= ZCC(2+42*(i-1),:)+ZCC(3+42*(i-1),:)+ZCC(4+42*(i-1),:)+ZCC(5+42*(i-1),:);
    X(3+9*(i-1),:)= ZCC(6+42*(i-1),:)+ZCC(7+42*(i-1),:)+ZCC(8+42*(i-1),:)+ZCC(9+42*(i-1),:)+ZCC(10+42*(i-1),:)+ZCC(11+42*(i-1),:)+ZCC(12+42*(i-1),:);
    X(4+9*(i-1),:)= ZCC(13+42*(i-1),:)+ZCC(14+42*(i-1),:)+ZCC(15+42*(i-1),:)+ZCC(16+42*(i-1),:)+ZCC(17+42*(i-1),:)+ZCC(18+42*(i-1),:)+ZCC(19+42*(i-1),:)+ZCC(20+42*(i-1),:)+ZCC(21+42*(i-1),:)+ZCC(22+42*(i-1),:)+ZCC(23+42*(i-1),:)+ZCC(24+42*(i-1),:);
    X(5+9*(i-1),:)= ZCC(25+42*(i-1),:)+ZCC(26+42*(i-1),:)+ZCC(27+42*(i-1),:);
    X(6+9*(i-1),:)= ZCC(28+42*(i-1),:);
    X(8+9*(i-1),:)= ZCC(31+42*(i-1),:)+ZCC(29+42*(i-1),:);
    X(7+9*(i-1),:)= ZCC(32+42*(i-1),:)+ZCC(30+42*(i-1),:);
    X(9+9*(i-1),:)= ZCC(33+42*(i-1),:)+ZCC(34+42*(i-1),:)+ZCC(35+42*(i-1),:)+ZCC(36+42*(i-1),:)+ZCC(37+42*(i-1),:)+ZCC(38+42*(i-1),:)+ZCC(39+42*(i-1),:)+ZCC(40+42*(i-1),:)+ZCC(41+42*(i-1),:)+ZCC(42+42*(i-1),:);

    Y(1+9*(i-1),:)= ZZSY(1+42*(i-1),:);
    Y(2+9*(i-1),:)= ZZSY(2+42*(i-1),:)+ZZSY(3+42*(i-1),:)+ZZSY(4+42*(i-1),:)+ZZSY(5+42*(i-1),:);
    Y(3+9*(i-1),:)= ZZSY(6+42*(i-1),:)+ZZSY(7+42*(i-1),:)+ZZSY(8+42*(i-1),:)+ZZSY(9+42*(i-1),:)+ZZSY(10+42*(i-1),:)+ZZSY(11+42*(i-1),:)+ZZSY(12+42*(i-1),:);
    Y(4+9*(i-1),:)= ZZSY(13+42*(i-1),:)+ZZSY(14+42*(i-1),:)+ZZSY(15+42*(i-1),:)+ZZSY(16+42*(i-1),:)+ZZSY(17+42*(i-1),:)+ZZSY(18+42*(i-1),:)+ZZSY(19+42*(i-1),:)+ZZSY(20+42*(i-1),:)+ZZSY(21+42*(i-1),:)+ZZSY(22+42*(i-1),:)+ZZSY(23+42*(i-1),:)+ZZSY(24+42*(i-1),:);
    Y(5+9*(i-1),:)= ZZSY(25+42*(i-1),:)+ZZSY(26+42*(i-1),:)+ZZSY(27+42*(i-1),:);
    Y(6+9*(i-1),:)= ZZSY(28+42*(i-1),:);
    Y(8+9*(i-1),:)= ZZSY(31+42*(i-1),:)+ZZSY(29+42*(i-1),:);
    Y(7+9*(i-1),:)= ZZSY(32+42*(i-1),:)+ZZSY(30+42*(i-1),:);
    Y(9+9*(i-1),:)= ZZSY(33+42*(i-1),:)+ZZSY(34+42*(i-1),:)+ZZSY(35+42*(i-1),:)+ZZSY(36+42*(i-1),:)+ZZSY(37+42*(i-1),:)+ZZSY(38+42*(i-1),:)+ZZSY(39+42*(i-1),:)+ZZSY(40+42*(i-1),:)+ZZSY(41+42*(i-1),:)+ZZSY(42+42*(i-1),:);
end

%计算直接消耗系数矩阵A
A = zeros(279,279);
for i = 1:279
    A(i,:)=biao_hang(i,:)/X(i,1);
end

%%处理碳数据
[Type Sheet Format]=xlsfinfo('Emission_Inventories_for_30_Provinces_2012.xlsx');
carbon_biao = zeros(47,1,30);
for i = 1:30
    carbon_biao(:,:,i) = xlsread('Emission_Inventories_for_30_Provinces_2012.xlsx',Sheet{i+1},'W5:W51');
end

%整合行
carbon_biao_tmp = zeros(9,1,30);
for i = 1:30
carbon_biao_tmp(1,:,i)=carbon_biao(1,:,i);
carbon_biao_tmp(2,:,i)=carbon_biao(2,:,i)+carbon_biao(3,:,i)+carbon_biao(4,:,i)+carbon_biao(5,:,i)+carbon_biao(6,:,i)+carbon_biao(7,:,i);
carbon_biao_tmp(3,:,i)=carbon_biao(9,:,i)+carbon_biao(10,:,i)+carbon_biao(11,:,i)+carbon_biao(12,:,i)+carbon_biao(13,:,i)+carbon_biao(14,:,i)+carbon_biao(15,:,i)+carbon_biao(8,:,i)+carbon_biao(16,:,i)+carbon_biao(17,:,i)+carbon_biao(18,:,i)+carbon_biao(19,:,i)+carbon_biao(20,:,i)+carbon_biao(21,:,i)+carbon_biao(22,:,i)+carbon_biao(23,:,i)+carbon_biao(24,:,i)+carbon_biao(25,:,i)+carbon_biao(26,:,i);
carbon_biao_tmp(4,:,i)=carbon_biao(27,:,i)+carbon_biao(28,:,i)+carbon_biao(29,:,i)+carbon_biao(30,:,i)+carbon_biao(31,:,i)+carbon_biao(32,:,i)+carbon_biao(33,:,i)+carbon_biao(34,:,i)+carbon_biao(35,:,i)+carbon_biao(36,:,i)+carbon_biao(37,:,i)+carbon_biao(38,:,i);
carbon_biao_tmp(5,:,i)=carbon_biao(39,:,i)+carbon_biao(40,:,i)+carbon_biao(41,:,i);
carbon_biao_tmp(6,:,i)=carbon_biao(42,:,i);
carbon_biao_tmp(7,:,i)=carbon_biao(43,:,i);
carbon_biao_tmp(8,:,i)=carbon_biao(44,:,i);
carbon_biao_tmp(9,:,i)=carbon_biao(45,:,i)+carbon_biao(46,:,i)+carbon_biao(47,:,i);
end

%还原处理行之后的表
carbon_biao_hang = zeros(9,30);
for i = 1:30
   carbon_biao_hang(:,i) = carbon_biao_tmp(:,:,i);
end

%%处理能源数据
[Type Sheet Format]=xlsfinfo('Province_Energy_Inventory_2012.xlsx');
energy_biao = zeros(47,20,30);
for i = 1:30
    energy_biao(:,:,i) = xlsread('Province_Energy_Inventory_2012.xlsx',Sheet{i+1},'B4:U50');
end

energy_biao_new = zeros(47,1,30);
for i = 1:30
    energy_biao_new(:,:,i)= energy_biao(:,1,i)*0.7143+energy_biao(:,2,i)*0.9+energy_biao(:,3,i)*0.2857+energy_biao(:,4,i)*0.6+energy_biao(:,5,i)*0.9714+energy_biao(:,6,i)*6.143+energy_biao(:,7,i)*3.5701+energy_biao(:,8,i)*1.3+energy_biao(:,9,i)*1.4286+energy_biao(:,10,i)*1.4714+energy_biao(:,11,i)*1.4714+energy_biao(:,12,i)*1.4571+energy_biao(:,13,i)*1.4286+energy_biao(:,14,i)*1.7143+energy_biao(:,15,i)*1.5714+energy_biao(:,16,i)*1.2+energy_biao(:,17,i)*12.143+energy_biao(:,18,i)*0.03412+energy_biao(:,19,i)*4.04+energy_biao(:,20,i);
end

%整合行
energy_biao_new_tmp = zeros(9,1,30);
for i = 1:30
energy_biao_new_tmp(1,:,i)=energy_biao_new(1,:,i);
energy_biao_new_tmp(2,:,i)=energy_biao_new(2,:,i)+energy_biao_new(3,:,i)+energy_biao_new(4,:,i)+energy_biao_new(5,:,i)+energy_biao_new(6,:,i)+energy_biao_new(7,:,i);
energy_biao_new_tmp(3,:,i)=energy_biao_new(9,:,i)+energy_biao_new(10,:,i)+energy_biao_new(11,:,i)+energy_biao_new(12,:,i)+energy_biao_new(13,:,i)+energy_biao_new(14,:,i)+energy_biao_new(15,:,i)+energy_biao_new(8,:,i)+energy_biao_new(16,:,i)+energy_biao_new(17,:,i)+energy_biao_new(18,:,i)+energy_biao_new(19,:,i)+energy_biao_new(20,:,i)+energy_biao_new(21,:,i)+energy_biao_new(22,:,i)+energy_biao_new(23,:,i)+energy_biao_new(24,:,i)+energy_biao_new(25,:,i)+energy_biao_new(26,:,i);
energy_biao_new_tmp(4,:,i)=energy_biao_new(27,:,i)+energy_biao_new(28,:,i)+energy_biao_new(29,:,i)+energy_biao_new(30,:,i)+energy_biao_new(31,:,i)+energy_biao_new(32,:,i)+energy_biao_new(33,:,i)+energy_biao_new(34,:,i)+energy_biao_new(35,:,i)+energy_biao_new(36,:,i)+energy_biao_new(37,:,i)+energy_biao_new(38,:,i);
energy_biao_new_tmp(5,:,i)=energy_biao_new(39,:,i)+energy_biao_new(40,:,i)+energy_biao_new(41,:,i);
energy_biao_new_tmp(6,:,i)=energy_biao_new(42,:,i);
energy_biao_new_tmp(7,:,i)=energy_biao_new(43,:,i);
energy_biao_new_tmp(8,:,i)=energy_biao_new(44,:,i);
energy_biao_new_tmp(9,:,i)=energy_biao_new(45,:,i)+energy_biao_new(46,:,i)+energy_biao_new(47,:,i);
end


%还原处理行之后的表
energy_biao_new_hang = zeros(9,30);
for i = 1:30
   energy_biao_new_hang(:,i) = energy_biao_new_tmp(:,:,i);
end

%%处理水数据
[Type Sheet Format]=xlsfinfo('Water-2012.xlsx');
water_biao = zeros(31,42);
water_biao = xlsread('Water-2012.xlsx',Sheet{2},'B3:AQ33');

%整合列
water_biao_tmp = zeros(31,9);
for i = 1:31
    water_biao_tmp(i,1)= water_biao(i,1);
    water_biao_tmp(i,2)= water_biao(i,2)+water_biao(i,3)+water_biao(i,4)+water_biao(i,5);
    water_biao_tmp(i,3)= water_biao(i,6)+water_biao(i,7)+water_biao(i,8)+water_biao(i,9)+water_biao(i,10)+water_biao(i,11)+water_biao(i,12);
    water_biao_tmp(i,4)= water_biao(i,13)+water_biao(i,14)+water_biao(i,15)+water_biao(i,16)+water_biao(i,17)+water_biao(i,18)+water_biao(i,19)+water_biao(i,20)+water_biao(i,21)+water_biao(i,22)+water_biao(i,23)+water_biao(i,24);
    water_biao_tmp(i,5)= water_biao(i,25)+water_biao(i,26)+water_biao(i,27);
    water_biao_tmp(i,6)= water_biao(i,28);
    water_biao_tmp(i,8)= water_biao(i,31)+water_biao(i,29);
    water_biao_tmp(i,7)= water_biao(i,32)+water_biao(i,30);
    water_biao_tmp(i,9)= water_biao(i,33)+water_biao(i,34)+water_biao(i,35)+water_biao(i,36)+water_biao(i,37)+water_biao(i,38)+water_biao(i,39)+water_biao(i,40)+water_biao(i,41)+water_biao(i,42);
end


%%计算各省各部门水能碳消耗系数
water_use_old = water_biao_tmp.';
water_use = zeros(279,1);
carbon_use = zeros(270,1);
energy_use = zeros(270,1);
for i = 1:31
    water_use(1+9*(i-1):9*i,:) = water_use_old(:,i);
end
for i = 1:30
    carbon_use(1+9*(i-1):9*i,:) = carbon_biao_hang(:,i);
    energy_use(1+9*(i-1):9*i,:) = energy_biao_new_hang(:,i);
end

water_xishu = zeros(279,279);
carbon_xishu = zeros(270,270);
energy_xishu = zeros(270,270);
for i = 1:279
    water_xishu(i,i)=water_use(i,:)/X(i,:);
end

for i = 1:25*9
    carbon_xishu(i,i)= carbon_use(i,:)/X(i,:);
    energy_xishu(i,i)= energy_use(i,:)/X(i,:);
end

for i = 25*9+1:270
    carbon_xishu(i,i)= carbon_use(i,:)/X(i+9,:);
    energy_xishu(i,i)= energy_use(i,:)/X(i+9,:);
end


water_xishu_sum = zeros(279,1);
carbon_xishu_sum = zeros(270,1);
energy_xishu_sum = zeros(270,1);
for i = 1:279
    water_xishu_sum(i,:)=water_xishu(i,i);
end

for i = 1:25*9
    carbon_xishu_sum(i,:)= carbon_xishu(i,i);
    energy_xishu_sum(i,:)= energy_xishu(i,i);
end

for i = 25*9+1:270
    carbon_xishu_sum(i,:)= carbon_xishu(i,i);
    energy_xishu_sum(i,:)= energy_xishu(i,i);
end

for i = 1:31
    for j = 1:9
        water_xishu_bumen(j,i) = water_xishu(j+9*(i-1),j+9*(i-1));
    end
end

for i = 1:30
    for j = 1:9
        carbon_xishu_bumen(j,i) = carbon_xishu(j+9*(i-1),j+9*(i-1));
        energy_xishu_bumen(j,i) = energy_xishu(j+9*(i-1),j+9*(i-1));
    end
end


%%总使用量
water_shiyong = zeros(9,1,31);
for i = 1:31
    water_shiyong(:,:,i) = water_use(1+9*(i-1):9*i,:);
end
carbon_shiyong = zeros(9,1,30);
energy_shiyong = zeros(9,1,30);
for i = 1:30
    carbon_shiyong(:,:,i) = carbon_use(1+9*(i-1):9*i,:);
    energy_shiyong(:,:,i) = energy_use(1+9*(i-1):9*i,:);
end


%%区域总足迹计算
%分部门
water_sum = zeros(9,31,31);
for i = 1:31
    for j = 1:31
        water_sum(:,j,i) = (water_xishu(1+9*(i-1):9*i,1+9*(i-1):9*i) * (sum(biao_hang(1+9*(i-1):9*i,1+9*(j-1):9*j),2)+sum(biao_Y_new(1+9*(i-1):9*i,1+5*(j-1):5*j),2)))/10000;
    end
end


water_local = zeros(9,31);
for i = 1:31
    water_local(:,i) = water_sum(:,i,i);
end

water_export = zeros(9,31);
for i = 1:31
    water_export(:,i) = sum(water_sum(:,:,i),2)-water_local(:,i);
end

water_extenral = zeros(9,31);
water_import = zeros(9,31);
for i = 1:31
   water_extenral(:,i) = sum(water_sum(:,i,:),3)-water_local(:,i);
   water_import(:,i) = sum(water_sum(:,i,:),3)-water_local(:,i);
end

carbon_sum = zeros(9,30,30);
energy_sum = zeros(9,30,30);
for i = 1:30
    for j = 1:30
        carbon_sum(:,j,i) = carbon_xishu(1+9*(i-1):9*i,1+9*(i-1):9*i) * (sum(biao_hang_carbon_energy(1+9*(i-1):9*i,1+9*(j-1):9*j),2)+sum(biao_Y_carbon_energy(1+9*(i-1):9*i,1+5*(j-1):5*j),2));
        energy_sum(:,j,i) = (energy_xishu(1+9*(i-1):9*i,1+9*(i-1):9*i) * (sum(biao_hang_carbon_energy(1+9*(i-1):9*i,1+9*(j-1):9*j),2)+sum(biao_Y_carbon_energy(1+9*(i-1):9*i,1+5*(j-1):5*j),2)))/1000;
    end
end

for i = 1:30
    carbon_local(:,i) = carbon_sum(:,i,i);
    energy_local(:,i) = energy_sum(:,i,i);
end

for i = 1:30
    carbon_export(:,i) = sum(carbon_sum(:,:,i),2)-carbon_local(:,i);
    energy_export(:,i) = sum(energy_sum(:,:,i),2)-energy_local(:,i);
end

for i = 1:30
   carbon_extenral(:,i) = sum(carbon_sum(:,i,:),3)-carbon_local(:,i);
   carbon_import(:,i) = sum(carbon_sum(:,i,:),3)-carbon_local(:,i);
   energy_extenral(:,i) = sum(energy_sum(:,i,:),3)-energy_local(:,i);
   energy_import(:,i) = (sum(energy_sum(:,i,:),3)-energy_local(:,i));
end

%不分部门
water_export_sum = sum(water_export);
water_extenral_sum = sum(water_extenral);
water_import_sum = sum(water_import);
carbon_export_sum = sum(carbon_export);
carbon_extenral_sum = sum(carbon_extenral);
carbon_import_sum = sum(carbon_import);
energy_export_sum = sum(energy_export);
energy_extenral_sum = sum(energy_extenral);
energy_import_sum = sum(energy_import);

for i = 1:31
    for j = 1:31
        tu_water_export_sheng(j+31*(i-1),:) = sum(water_sum(:,j,i));
    end
end

for i = 1:30
    for j = 1:30
        tu_carbon_export_sheng(j+30*(i-1),:) = sum(carbon_sum(:,j,i));
    end
end

for i = 1:30
    for j = 1:30
        tu_energy_export_sheng(j+30*(i-1),:) = sum(energy_sum(:,j,i));
    end
end

tu_water_sum_xiaofei = sum(sum(water_sum,3)).';
tu_carbon_sum_xiaofei = (sum(sum(carbon_sum,3))).';
tu_energy_sum_xiaofei = sum(sum(energy_sum,3)).';