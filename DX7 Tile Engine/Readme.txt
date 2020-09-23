#################################
########## Map Format ###########
#################################
	
	Map Header:

	16 bits - Width in tiles
	16 bits - Height in tiles
	16 bits - Starting X position in tiles
	16 bits - Starting Y position in tiles
	8 bits  - Tileset width in tiles
	8 bits  - Tileset height in tiles
	8 bits  - Tileset tile width in pixels
	8 bits  - Tileset tile height in pixels
	
        
	Array of tile structs:       	

            	1 Element: 

		8 bits - Graphic index of the current tile (according to tileset)
		8 bits - Whether or not the tile is walkable (1 indicates yes, 0 indicates no)
        
	Map Footer:

        16 bits - Number of portals
        
	Array of portal structs:
	
		1 Element:        

            	16 bits  - X position in tiles
		16 bits  - Y position in tiles
		16 bits  - Destination X in tiles
		16 bits  - Destination Y in tiles
		256 bits - Destination map name


#################################
########### Debug Mode ##########
#################################

To switch between debug mode and normal mode, change the second line (the one after Option Explicit) in modEngine.

#Const DebugMode = <Either True or False>


#################################
############ Credits ############
#################################

Thanks to Lucky (http://www.rookscape.com/vbgaming/) for some of the tiles and a few of the ideas (and the odd piece of code).

Thanks to the author of DXRE (Search on PSC) for the character graphics (which were reassembled into one file by myself).

Thanks to http://www.directx4vb.com/ for the excellent DirectX tutorials