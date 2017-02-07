# AutoMine - Minesweeper X Solver
# Copyright (c) 2017 HizkiFW

import win32gui, win32api, win32con
import time
import signal, os, random

tileSize = 16
width = 29  # TODO: auto detect
height = 15 # tiles w and h minus 1

# Easy: 7,7
# Intermediate: 15,15
# Expert: 29,15

marginTop = 100
marginLeft = 15
buttonTop = 72
colors = (0xEEEEEE, 0x2196F3, 0x4CAF50, 0xF44336, 0x3F51B5, 0xD50000, 0x009688, 0x9C27B0, 0x212121, 0xED1C24, 0xE0E0E0)

# Get window info
msxWin = win32gui.FindWindow(None, "Minesweeper X")
winRect = win32gui.GetWindowRect(msxWin)
cliRect = win32gui.GetClientRect(msxWin)
rect = (winRect[0], winRect[1], cliRect[2], cliRect[3])

around = ((0, 1), (0, -1), (1, 1), (1, 0), (1, -1), (-1, 1), (-1, 0), (-1, -1))
tilecache = {}

# Mouse functions
def lclick(x, y):
	win32api.SetCursorPos((x, y))
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
	time.sleep(0.01)
def rclick(x, y):
	win32api.SetCursorPos((x, y))
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)
	time.sleep(0.01)
def dclick(x, y):
	win32api.SetCursorPos((x, y))
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)
	time.sleep(0.01)

def getPixel(i_x, i_y):
	i_desktop_window_id = win32gui.GetDesktopWindow()
	i_desktop_window_dc = win32gui.GetWindowDC(i_desktop_window_id)
	long_colour = win32gui.GetPixel(i_desktop_window_dc, i_x, i_y)
	i_colour = int(long_colour)
	#return (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)
	return ((i_colour & 0xff) << 16) + (i_colour & 0xff00) + ((i_colour & 0xff0000) >> 16)
	#return i_colour

# Game functions
def gameRestart():
	global tilecache
	tilecache = {}
	lclick(rect[0] + (rect[2] / 2), rect[1] + buttonTop)

def gameStatus():
	color = getPixel(rect[0] + (rect[2] / 2), rect[1] + buttonTop)
	
	if color == 0x00FF00:
		return 0 # normal
	elif color == 0xFF0000:
		return 2 # lost
	else:
		return 1 # win?

# Tile functions
def tileCoord(x, y):
	# Converts tile X and Y to screen coordinates
	return rect[0] + marginLeft + tileSize*x + tileSize/2, rect[1] + marginTop + tileSize*y + tileSize/2

def tileStatus(x, y):
	# Gets tile status
	# 0 = empty, 1-8 = n bombs around, 9 = flagged, 10 = new/unopened
	global tilecache
	try:
		return tilecache[x, y]
	except:
		rx, ry = tileCoord(x, y)
		color = getPixel(rx, ry)
		for i in range(len(colors)):
			if colors[i] == color:
				if i < 10:
					tilecache[x, y] = i
				return i
		return 99

def tileAround(x, y, type):
	n = 0
	for offset in around:
		if tileStatus(x + offset[0], y + offset[1]) == type: n += 1
	return n

def tileOpen(x, y):
	rx, ry = tileCoord(x, y)
	lclick(rx, ry)

def tileOpenCorners():
	tileOpen(0, 0)
	tileOpen(0, height)
	tileOpen(width, 0)
	tileOpen(width, height)

def tileFlag(x, y):
	rx, ry = tileCoord(x, y)
	rclick(rx, ry)
	
def tileExpose(x, y):
	rx, ry = tileCoord(x, y)
	dclick(rx, ry)

# Tile analysis
def analyzeAround(x, y, dbg):
	for offset in around:
		t = tileStatus(x + offset[0], y + offset[1])
		if t > 0 and t < 9:
			analyzeTile(x + offset[0], y + offset[1], dbg+1)

def analyzeTile(x, y, dbg):
	s = tileStatus(x, y)
	if gameStatus() == 2 or dbg > 20 or s == 0 or s == 9 or s == 10:
		return 0
	else:
		# Main logic part
		taflag = tileAround(x, y, 9)
		tanew = tileAround(x, y, 10)
		if(taflag + tanew) == s and not (s == taflag):
			for offset in around:
				if tileStatus(x + offset[0], y + offset[1]) == 10:
					print (" "*dbg) + "Flag! {" + str(s) + ")"
					tileFlag(x + offset[0], y + offset[1])
					analyzeAround(x, y, dbg)
					return 1
		
		if taflag == s and tanew > 0:
			print (" "*dbg) + "Expose! (" + str(s) + ")"
			tileExpose(x, y)
			analyzeAround(x, y, dbg)
			return 1
		
		return 0

def guessAroundTile(x, y):
	for offset in around:
		if tileStatus(x + offset[0], y + offset[1]) == 10:
			tileOpen(x + offset[0], y + offset[1])
			return

def init():
	print "Starting AutoMine"
	global tilecache
	tilecache = {}
	
	# Open up corners
	while 1:
		tileOpenCorners()
		if gameStatus() == 0:
			break
		else:
			time.sleep(0.5)
			gameRestart()
	
	x = 0
	y = 0
	guess = 0
	# Tile scan loop
	while 1:
		# Exit if won or lost
		if gameStatus() > 0: return
		
		print (x, y)
		
		# Set cursor position while scanning
		rx, ry = tileCoord(x, y)
		win32api.SetCursorPos((rx, ry))
		
		# Guess if no changes after 2 scan cycles
		if guess > 1 and tileStatus(x, y) < 9 and tileAround(x, y, 10) > 0:
			print "Guess!"
			guess = 0
			
			guessAroundTile(x, y)
			
			# Restart if guess fails
			if gameStatus() == 2:
				print "EXPLOSION!"
				time.sleep(1)
				init()
		
		# Analyze tile, reset guess counter if something happened (flagged, opened tiles)
		if analyzeTile(x, y, 0) > 0:
			guess = 0

		# Increment X and Y
		x += 1
		if x > width:
			x = 0
			y += 1
		if y > height:
			x = 0
			y = 0
			guess += 1

# Initialize
if __name__ == "__main__":
	init()
