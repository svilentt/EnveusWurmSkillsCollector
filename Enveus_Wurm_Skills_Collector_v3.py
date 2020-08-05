#local
import dumpsLocationHandler
import excelSheetUpdater

print('Enveus Wurm Skills Collector v3 by Sveny\n')

skills = {
	'Body' : '1,0',
	'Body control' : '1,0',
	'Body stamina' : '1,0',
	'Body strength' : '1,0',
	'Mind' : '1,0',
	'Mind logic' : '1,0',
	'Mind speed' : '1,0',
	'Soul' : '1,0',
	'Soul depth' : '1,0',
	'Soul strength' : '1,0',
	'Alchemy' : '1,0',
	'Natural substances' : '1,0',
	'Archery' : '1,0',
	'Long bow' : '1,0',
	'Medium bow' : '1,0',
	'Short bow' : '1,0',
	'Axes' : '1,0',
	'Hatchet' : '1,0',
	'Huge axe' : '1,0',
	'Large axe' : '1,0',
	'Small axe' : '1,0',
	'Carpentry' : '1,0',
	'Bowyery' : '1,0',
	'Fine carpentry' : '1,0',
	'Fletching' : '1,0',
	'Ship building' : '1,0',
	'Toy making' : '1,0',
	'Climbing' : '1,0',
	'Clubs' : '1,0',
	'Huge club' : '1,0',
	'Coal-making' : '1,0',
	'Cooking' : '1,0',
	'Baking' : '1,0',
	'Beverages' : '1,0',
	'Butchering' : '1,0',
	'Dairy food making' : '1,0',
	'Hot food cooking' : '1,0',
	'Digging' : '1,0',
	'Fighting' : '1,0',
	'Aggressive fighting' : '1,0',
	'Defensive fighting' : '1,0',
	'Normal fighting' : '1,0',
	'Shield bashing' : '1,0',
	'Taunting' : '1,0',
	'Weaponless fighting' : '1,0',
	'Healing' : '1,0',
	'First aid' : '1,0',
	'Knives' : '1,0',
	'Butchering knife' : '1,0',
	'Carving knife' : '1,0',
	'Masonry' : '1,0',
	'Stone cutting' : '1,0',
	'Mauls' : '1,0',
	'Large maul' : '1,0',
	'Medium maul' : '1,0',
	'Small maul' : '1,0',
	'Milling' : '1,0',
	'Mining' : '1,0',
	'Nature' : '1,0',
	'Animal husbandry' : '1,0',
	'Animal taming' : '1,0',
	'Botanizing' : '1,0',
	'Farming' : '1,0',
	'Fishing' : '1,0',
	'Foraging' : '1,0',
	'Forestry' : '1,0',
	'Gardening' : '1,0',
	'Meditating' : '1,0',
	'Milking' : '1,0',
	'Papyrusmaking' : '1,0',
	'Paving' : '1,0',
	'Pottery' : '1,0',
	'Prospecting' : '1,0',
	'Religion' : '1,0',
	'Artifacts' : '1,0',
	'Channeling' : '1,0',
	'Exorcism' : '1,0',
	'Prayer' : '1,0',
	'Preaching' : '1,0',
	'Ropemaking' : '1,0',
	'Shields' : '1,0',
	'Large metal shield' : '1,0',
	'Large wooden shield' : '1,0',
	'Medium metal shield' : '1,0',
	'Medium wooden shield' : '1,0',
	'Small metal shield' : '1,0',
	'Small wooden shield' : '1,0',
	'Smithing' : '1,0',
	'Blacksmithing' : '1,0',
	'Jewelry smithing' : '1,0',
	'Locksmithing' : '1,0',
	'Metallurgy' : '1,0',
	'Armour smithing' : '1,0',
	'Chain armour smithing' : '1,0',
	'Plate armour smithing' : '1,0',
	'Shield smithing' : '1,0',
	'Weapon smithing' : '1,0',
	'Blades smithing' : '1,0',
	'Weapon heads smithing' : '1,0',
	'Swords' : '1,0',
	'Longsword' : '1,0',
	'Shortsword' : '1,0',
	'Two handed sword' : '1,0',
	'Tailoring' : '1,0',
	'Cloth tailoring' : '1,0',
	'Leatherworking' : '1,0',
	'Thatching' : '1,0',
	'Tracking' : '1,0',
	'Thievery' : '1,0',
	'Lock picking' : '1,0',
	'Stealing' : '1,0',
	'Traps' : '1,0',
	'War machines' : '1,0',
	'Catapults' : '1,0',
	'Ballistae' : '1,0',
	'Trebuchet' : '1,0',
	'Turrents' : '1,0',
	'Woodcutting' : '1,0'
}

def extractPlayerNameFromFile(fullFilePath):
	with open(fullFilePath) as f:
		fileLines = f.readlines()
		return fileLines[1].split()[1]
		
def extractSkillsFromFile(fullFilePath):
	with open(fullFilePath) as f:
		fileLines = f.readlines()
		for line in fileLines[5:]:
			skill = line.split(':')
			skillNameParts = skill[0].split(' ')
			skillName = ""
			for skillNamePart in skillNameParts:
				if skillNamePart != '':
					skillName += skillNamePart + ' '
			skillName = skillName[:-1]
			skillValueParts = skill[1].split()[0].split('.')
			skillValueParts[1] += '0';
			skillValue = skillValueParts[0] + ',' + skillValueParts[1][0:2]
			if skillName in skills:
				if skillValueParts[0] != '0':
					skills[skillName] = skillValue
	return list(skills.values())

if not dumpsLocationHandler.isFolderSet() or not dumpsLocationHandler.isFolderCorrect():
	print('Asking for dumps folder location...')
	dumpsLocationHandler.askForFolder()

if not dumpsLocationHandler.isFolderCorrect():
	print('The selected folder is not correct or empty. Please restart the application and select a folder again.')
	print("Press enter to exit the console...")
	input()
else:	
	print('Getting the latest dump...')				
	latestDump = dumpsLocationHandler.getLatestDumpFile()

	print('Fetching data...')
	playerName = extractPlayerNameFromFile(latestDump)
	playerData = extractSkillsFromFile(latestDump)

	print('Updating remote spreadsheet...')
	if excelSheetUpdater.updateSheet(playerName, playerData):
		print('Update succeeded!')
	
	print("Press enter to exit the console...")
	input()	