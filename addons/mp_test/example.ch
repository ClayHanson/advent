start]msgbox "A start event. It's optional, but if it is enabled, it will act like an event trigger. The code executes when the player first loads the map."
0]./maps/start
1]./maps/start{_x15_y15}
9E]msgbox "An event trigger. Format: (Trigger #)E](WSCRIPT Code)"&vbCr&"Code example: 9E]msgbox ""test"" : msgbox ""test2"""
à]NPC 1. SHOWS OFF PLAYER VARIABLES. NAME: %NAME% MONEY: %MONEY% AGE: %AGE%
á]NPC 2. GIVES FREE ITEM.{ItemName}[ItemDescription]
â]NPC 3. GIVES ITEM THAT COSTS $15.{ItemCostName$0015}[ItemDescription]
ã]NPC 4. GIVES ITEM THAT COSTS $15 AND REQUIRES THE PLAYER'S AGE TO BE OVER 18.{ItemCostAgeName$0015%018}[ItemDescription]
ä]NPC 5. SHOWS\nOFF\nNEW\nLINES.
b]a doesn't have a definition, but b does, so b is changed to an npc.
$]150