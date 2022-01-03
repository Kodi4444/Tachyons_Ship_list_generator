import os
import xlsxwriter

exclude_Star_Whale = True
exclude_Unity_Guard_Cruiser = True
exclude_Slaver_guard_ship = True
exclude_Slaver_escort_ship = True

Trading_Stations = []
Asteroid = []
Star_Whale = []
Unity_Trader_ship = []
table = []
temp = []

for subdir, dirs, files in os.walk("server\\DataBase\\Ships"):
    # print(subdir)
    path = subdir.split('\\')
    # print(path)
    ship_id = path[-1]
    # print(ship_id)
    ship = []
    ship.append(ship_id)
    for file in files:
        # print(file)
        ship.append(file)
        with open(subdir + '\\' + file, "r") as myfile:
            data = myfile.read().splitlines()
            # print(data)
            ship.append(data)
    # print(ship)
    for e in ship:
        if len(e) > 1 and isinstance(e, list):
            for i in e:
                print(end='    ')
                print(i)
        else:
            print(e)
    print('test'.center(100, "#"))
    # print(ship)

    if len(ship) > 1:
        print(ship[4][3])
        if ship[4][3] == 'NAME=Trading Station':
            loc_id = [ship[12][-1], ship_id]
            Trading_Stations.append(loc_id)
        if ship[4][3] == 'NAME=Asteroid':
            loc_id = [ship[12][-1], ship_id]
            Asteroid.append(loc_id)
        if not exclude_Star_Whale:
            if ship[4][3] == 'NAME=Star Whale':
                loc_id = [ship[12][-1], ship_id]
                Star_Whale.append(loc_id)
        if ship[4][3] == 'NAME=Unity Trader ship':
            loc_id = [ship[12][-1], ship_id]
            Unity_Trader_ship.append(loc_id)

        ship_name_raw = ship[4][3]
        ship_name = ship_name_raw[5:]
        REPAIR = False
        for i in ship[10]:
            if 'REPAIR' in i:
                REPAIR = True

        SYS_SHOP = False
        for i in ship[10]:
            if 'SYS_SHOP' in i:
                SYS_SHOP = True

        SHIPS_SHOP = False
        for i in ship[10]:
            if 'SHIPS_SHOP' in i:
                SHIPS_SHOP = True

        MIS_SHOP = False
        for i in ship[10]:
            if 'MIS_SHOP' in i:
                MIS_SHOP = True

        DRONE_SHOP = False
        for i in ship[10]:
            if 'DRONE_SHOP' in i:
                DRONE_SHOP = True
        if exclude_Star_Whale and ship_name == "Star Whale":
            pass
        elif exclude_Unity_Guard_Cruiser and ship_name == "Unity Guard Cruiser":
            pass
        elif exclude_Slaver_guard_ship and ship_name == "Slaver guard ship":
            pass
        elif exclude_Slaver_escort_ship and ship_name == "Slaver escort ship":
            pass
        else:
            temp = [ship_name, ship[12][-1], int(REPAIR), int(SYS_SHOP), int(SHIPS_SHOP), int(MIS_SHOP),
                    int(DRONE_SHOP), ship_id]
            table.append(temp)
        print('temp =', temp)

    print('test2'.center(100, "#"))
    print()


print('Trading_Stations')
print(Trading_Stations)
print(len(Trading_Stations))

print('Asteroid')
print(Asteroid)
print(len(Asteroid))

print('Star_Whale')
print(Star_Whale)
print(len(Star_Whale))

print('Unity_Trader_ship')
print(Unity_Trader_ship)
print(len(Unity_Trader_ship))

# print(ship)
# print(ship.index('systems.dat'))
# print(ship[4].index('NAME=Trading Station'))
# print(ship[10])
# for a in ship[10]:
#     if 'REPAIR' in a:
#         print(True, '$$$$$$$$$$$$$$$$$$$$$')
# print(type(ship[4][3]))


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('ships.xlsx')
worksheet = workbook.add_worksheet()


worksheet.write(0, 0, 'Name')
worksheet.write(0, 1, 'Location')
worksheet.write(0, 2, 'Repair')
worksheet.write(0, 3, 'SYS_SHOP')
worksheet.write(0, 4, 'SHIPS_SHOP')
worksheet.write(0, 5, 'MIS_SHOP')
worksheet.write(0, 6, 'DRONE_SHOP')
worksheet.write(0, 7, 'Ship ID')


row = 1
col = 0

# Iterate over the data and write it out row by row.
for s_name, s_location, s_repair, s_SYS_SHOP, s_SHIPS_SHOP, s_MIS_SHOP, s_DRONE_SHOP, s_id in table:
    worksheet.write(row, col,     s_name)
    worksheet.write(row, col + 1, s_location)
    worksheet.write(row, col + 2, s_repair)
    worksheet.write(row, col + 3, s_SYS_SHOP)
    worksheet.write(row, col + 4, s_SHIPS_SHOP)
    worksheet.write(row, col + 5, s_MIS_SHOP)
    worksheet.write(row, col + 6, s_DRONE_SHOP)
    worksheet.write(row, col + 7, s_id)

    row += 1

workbook.close()
