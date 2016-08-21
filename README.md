# LanParty
Chat client and game launcher designed for simplifying LAN parties. Written in Visual Basic 6.

To do:

Doesn't currently compile. I'm in the process of adding multi-user private chats, but as it was an afterthought it won't be pretty.

The LanParty Client is designed to be distributed with a pre-made database of games of your choosing, ensuring that each game is available on the end user's computer.

There is some basic encryption for each packet based on a hard-coded key, and each user will generate their own unique key for each new user they connect to.

The program uses UDP for discovery and communication.
