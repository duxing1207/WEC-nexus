clear all;
tic;

%%处理MRIO数据
%读整体
[Type Sheet Format]=xlsfinfo('WEC nexus data_2030.xlsx');
biao = zeros(1302,1302);
biao = xlsread('WEC nexus data_2030.xlsx','Output_Raw','C4:AXD1305');
biao_Y = zeros(1302,155);
biao_Y = xlsread('WEC nexus data_2030.xlsx','Output_Raw','AXE4:BDC1305');

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

biao_hang_future(1:9*25,1:9*25) = biao_hang(1:9*25,1:9*25);
biao_hang_future(1:9*25,9*25+1:30*9) = biao_hang(1:9*25,9*26+1:31*9);
biao_hang_future(9*25+1:30*9,1:9*25) = biao_hang(9*26+1:31*9,1:9*25);
biao_hang_future(9*25+1:30*9,9*25+1:30*9) = biao_hang(9*26+1:31*9,9*26+1:31*9);
biao_Y_future(1:9*25,1:5*25) = biao_Y_new(1:9*25,1:5*25);
biao_Y_future(1:9*25,5*25+1:30*5) = biao_Y_new(1:9*25,5*26+1:31*5);
biao_Y_future(9*25+1:30*9,1:5*25) = biao_Y_new(9*26+1:31*9,1:5*25);
biao_Y_future(9*25+1:30*9,5*25+1:30*5) = biao_Y_new(9*26+1:31*9,5*26+1:31*5);

water_future = xlsread('WEC nexus data_2030.xlsx','WEC use amount','C2:C31');
carbon_future = xlsread('WEC nexus data_2030.xlsx','WEC use amount','D2:D31');
energy_future = xlsread('WEC nexus data_2030.xlsx','WEC use amount','E2:E31');
money_xishu_future = xlsread('WEC nexus data_2030.xlsx','GDP growth rate_2018 to 2035','T2:T31');



for i = 1:30
    for j = 1:30
        biao_future_sum(i,j) = sum(sum(biao_hang_future(1+9*(i-1):9*i,1+9*(j-1):9*j),2)+sum(biao_Y_future(1+9*(i-1):9*i,1+5*(j-1):5*j),2));
    end
end

money_sum_future_1 = sum(biao_future_sum,2);
for i = 1:30
    money_sum_future(i,1) = money_sum_future_1(i,1)* money_xishu_future(i,1);
end


for k = 1:30
    water_xishu_future(k,:) = water_future(k,:)/money_sum_future(k,:);
    carbon_xishu_future(k,:)= carbon_future(k,:)/money_sum_future(k,:);
    energy_xishu_future(k,:)= energy_future(k,:)/money_sum_future(k,:);
end

for i = 1:30
    for j = 1:30
        water_sum_future(i,j) = water_xishu_future(i,1) * biao_future_sum(i,j);
        carbon_sum_future(i,j) = carbon_xishu_future(i,1) * biao_future_sum(i,j);
        energy_sum_future(i,j) = energy_xishu_future(i,1) * biao_future_sum(i,j);
    end
end

tu_water_sum_comsumption_future = sum(water_sum_future).';
tu_carbon_sum_comsumption_future = sum(carbon_sum_future).';
tu_energy_sum_comsumption_future = sum(energy_sum_future).';

for i = 1:30
    carbon_local_future(i,i) = carbon_sum_future(i,i);
    energy_local_future(i,i) = energy_sum_future(i,i);
    water_local_future(i,i) = water_sum_future(i,i);
end

carbon_export_future = carbon_sum_future-carbon_local_future;
energy_export_future = energy_sum_future-energy_local_future;
water_export_future = water_sum_future-water_local_future;
