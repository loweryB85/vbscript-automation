Dim objNetwork, strRemoteShare
Set objNetwork = WScript.CreateObject("WScript.Network")

strRemoteShare = "<network path>"

objNetwork.MapNetworkDrive "R:", strRemoteShare, True