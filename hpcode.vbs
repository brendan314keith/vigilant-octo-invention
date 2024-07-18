'translated from unix
main()

Sub main
	arg_count = wscript.arguments.count

	If wscript.arguments(0) = "-ver" Then
		Set winsock = Wscript.createobject("MSWINSOCK.Winsock", "winsock ")
		wscript.echo winsock.version
		Exit Sub
	End If

	If wscript.arguments(0) = "-kvb" Then
		Set winsock = Wscript.createobject("kvbwinsocklib.Winsock", "winsock ")
		wscript.echo winsock.version
		Exit Sub
	End If

	If arg_count < 3 Then 
		MsgBox "Usage: hpcode <iterations> <delay(sec)> <ipaddr> [port] [" & Chr(34) & "Custom message" & Chr(34) & "]" & vbcrlf & _
		" - use iterations = 0 for infinite loop" & vbcrlf & _
		" - use delay = 0 to reset printer message" & vbcrlf & _
		" - <iterations> is ignored when using a custom message" & vbcrlf & _
		" - " & Chr(34) & "Custom message" & Chr(34) & " must have double quotes"
		
		Exit Sub
	Else 
		use_custom_message = vbfalse
		
		looptimes = Int(wscript.arguments(0))
		mydelay = Int(wscript.arguments(1))
		myprinter = wscript.arguments(2) 
		myport = 9100 
		If arg_count > 3 Then
			If IsNumeric(wscript.arguments(3)) Then
				myport = wscript.arguments(3) 
				cms=4
			Else 
				myport = 9100 
				cms=3
			End if 
		
			If arg_count > cms Then
				custom_message = wscript.arguments(cms)
				wscript.echo "Custom:" & custom_message 
				use_custom_message = vbtrue
			End If
		End If
	End If

	mydefault_message = "READY."

	If use_custom_message = vbtrue Then
		send_to_printer myprinter, myport, custom_message
		If mydelay > 0 Then
			wscript.sleep mydelay*1000
			send_to_printer myprinter, myport, mydefault_message
		End If
		Exit Sub
	End If
			
	If mydelay = 0 Then 
		wscript.echo "Resetting"
		send_to_printer myprinter, myport, mydefault_message
		Exit Sub
	End If

		
	Randomize
	loopsdone = 0
	
	Do 
		mymessage = worksafemessage()
'		mymessage = "READY."
		' Send the requested message...
		
		send_to_printer myprinter, myport, mymessage
		loopsdone = loopsdone + 1
		wscript.echo "|" & loopsdone & "| of |" & looptimes & "|" & " - " & mymessage
	
		wscript.sleep mydelay*1000
		
	Loop Until loopsdone = looptimes
	
	send_to_printer myprinter, myport, mydefault_message

End Sub


' Pretty standard socket stuff here - see the docs for Socket.pm for specifics.
' See also the PJL reference, _ available at http://hpdevelopersolutions.com.  
' Be prepared for a fascist registration process.  Nasty page.

sub send_to_printer(myprinter, myport, mymessage)
  
	Set winsock = Wscript.createobject("MSWINSOCK.Winsock", "winsock ")
	winsock.Remotehost = myprinter  
	winsock.RemotePort = myport   
	winsock.connect      'Make the connection

	While winsock.state <> 7 And winsock.state <> 8 And winsock.state <> 9 And winsock.state <> 0 And secs <> 25
	  Wscript.Sleep 200  'Sleep for a sec then check again
	  secs = secs + 1
	Wend

	'wscript.echo "Winsock state: " & winsock.state & vbcrlf & "Connection error:" & err.number & " - " & err.description
	If winsock.state = 7 Then
		cmdstr = Chr(27) & "%-12345X@PJL RDYMSG DISPLAY="  & Chr(34) & mymessage & Chr(34) & Chr(13) & Chr(10)
		winsock.senddata cmdstr
		'wscript.echo "Sent: " & cmdstr & " to " & myprinter & ":" & myport
		
		Wscript.Sleep 100  'Sleep for a sec then check again
		
		'winsock.senddata Chr(33) & "%-12345X" & Chr(13)
		winsock.close
	Else
		wscript.echo "Couldn't connect to printer: " & myprinter & ":" & myport
	End If
	
	'[ESC]%-12345X@PJL JOB
	'@PJL RDYMSG DISPLAY="Any message"
	'@PJL EOJ
	'[ESC]%-12345X
	

	'myprinter_address = inet_aton(myprinter) or die "inet_aton failed.  I think I also broke your 's' key.\n";
	'$addr = sockaddr_in(myport, _ myprinter_address) or die "sockaddr_in failed.  Don't you know where your sock_drawer is?\n";
	'$proto = getprotobyname('tcp') or die "getprotobyname failed.  That's probably a very bad sign.\n";
	'socket(S, _ AF_INET, _ SOCK_STREAM, _ $proto) or die "Socket creation failed.  Fire and brimstone await you.\n";
	'connect(S, _$addr) or die "Connection to myprinter failed.  I suggest empathy classes.\n"; 
	'select S; $| = 1; select STDOUT;
	'print S "\033%-12345X\@PJL RDYMSG DISPLAY =\"$message\"\n";
	'print S "\033%-12345X\n";
	'close S;

End Sub

Function message()
	allmessages = Array( _
	    "LETTER E IS LOW                 ", _
	    "INSERT 25 CENTS PER PAGE        ", _
	    "SPELL CHECK     MODE ON         ", _
	    "PRINTER         ON FIRE         ", _
	    "XRAY LEAK ALARM                 ", _
	    "404 DOCUMENT NOTFOUND           ", _
	    "CRUMB TRAY FULL                 ", _
	    "CORE COLLAPSE   IMMINENT        ", _
	    "41.4 JOB DENIED OUT OF MAYO     ", _
	    "INSERT FRESH    HAMSTER         ", _
	    "WARP CORE BREACH                ", _
	    "PING TIMEOUT                    ", _
	    "TONER HIGH                      ", _
	    "RANDOMIZING THE TONER           ", _
	    "NO CARRIER                      ", _
	    "LASER IS EMPTY  PLEASE REFILL   ", _
	    "WHITE TONER LOW                 ", _
	    "PRINTING COPY   1 OF 999999     ", _
	    "BACK IN 15      MINUTES         ", _
	    "EXACT CHANGE    ONLY            ", _
	    "PAPER CUT       INSERT BAND-AID ", _
	    "YOUR COMMERCIAL MESSAGE HERE    ", _
	    "PLEASE HOLD FOR THE OPERATOR    ", _
	    "NOW BOOTING     WINDOWS XP      ", _
	    "PLE  E RE L  N     AS    A IG   " )
	
	message =    allmessages(Int(Rnd * UBound(allmessages)) + 1)

End Function


Function worksafemessage()
	allmessages = Array( _
	    "PRINTER LOW ON  CHEESE", _
	    "GREETINGS       EARTHLINGS", _
	    "LETTER E IS LOW", _
	    "INSERT 25 CENTS FOR NEXT 5 MINS", _
	    "PRINTER ON FIRE JUST KIDDING", _
	    "PAPER TOO SPICY INSERT TUMS", _
	    "SPELL CHECK     MODE ON", _
	    "NO SMOKING,     PLEASE", _
	    "SAY PLEASE", _
	    "I LOVE THE      FAX MACHINE", _
	    "SURELY YOU JEST", _
	    "I'M A LITTLE    TEAPOT", _
	    "LEARN TO USE A  DAMN PENCIL", _
	    "LEARN TO USE    A PENCIL", _
	    "NOBODY PRINTS   ANYTHING FUN", _
	    "I WISH I WAS A  FAX MACHINE", _
	    "PLEASE PRINT    BETTER STUFF", _
	    "I WISH I HAD    ARMS AND LEGS", _
	    "YOU DON'T REALLYWANT TO PRINT", _
	    "GO BOTHER THE   COPIER", _
	    "I WISH I COULD  SING", _
	    "I DON'T LIKE YOUEITHER", _
	    "I SEE DEAD      PEOPLE", _
	    "DON'T LOOK AT MELIKE I OWE YOU", _
	    "GOODBYE CRUEL   WORLD", _
	    "THE COPIER IS   MOCKING ME", _
	    "DON'T GET YOUR  HOPES UP", _
	    "I WAS A TOASTER IN MY LAST LIFE", _
	    "DOCUMENT NEEDS  MORE SALT", _
	    "YOU HAVE NO IDEAHOW BORED I AM", _
	    "THAT OUTFIT DOESNOT SUIT YOU", _
	    "HOW I LONG FOR AHOT BATH", _
	    "I WISH I HAD    A BOOK TO READ", _
	    "SPELLING IS NOT YOUR FORTE", _
	    "YOU'D HAVE TO   BEAT ME FIRST", _
	    "WRITING IS NOT YOUR FORTE", _
	    "DOCUMENT DELETEDWHAT NEXT?", _
	    "THIS ISN'T WHAT I STUDIED FOR", _
	    "TIME TO FEED THEMONKEY", _
	    "I WANTED TO BE ANICE VCR", _
	    "PENGUIN ON FIRE", _
	    "HEY, NICE SHOES", _
	    "FEED ME", _
	    "ON FIRE", _
	    "BANANA STUCK", _
	    "XRAY LEAK ALARM", _
	    "STARTING JAVA...", _
	    "404 DOCUMENT NOTFOUND", _
	    "CALL 911", _
	    "PORN TRAY EMPTY", _
	    "CRUMB TRAY EMPTY", _
	    "CORE COLLAPSE   IMMINENT", _
	    "41.4 JOB DENIED OUT OF MAYO", _
	    "TYPE IT YOURSELF", _
	    "KERNEL PANIC", _
	    "JOB TOO LONG", _
	    "YOUR JOB IS MINE", _
	    "TAKE A TPYING   CLASS ALREADY", _
	    "I'M TIRED", _
	    "HAMSTER DEAD    TURN WHEEL", _
	    "INSERT FRESH    HAMSTER", _
	    "WHEN HELL       FREEZES OVER", _
	    "WARP CORE BREACH", _
	    "CARESS ME AND   I AM YOURS", _
	    "PING TIMEOUT", _
	    "REMOVE PAPER", _
	    "DRY CLEAN CYCLE", _
	    "HAND WASH ONLY", _
	    "TONER HIGH", _
	    "SUPER UNLEADED", _
	    "LOW ON CAFFEINE", _
	    "NO SMOKING", _
	    "I HATE MY JOB", _
	    "RANDOMIZING THE TONER", _
	    "NO CARRIER", _
	    "GIRAFFE NEEDS   MORE LEMONS", _
	    "EVERYTHING IS   BEAUTIFUL", _
	    "I CAN'T FEEL MY LEGS", _
	    "EVERYONE THINKS YOU'RE INSANE", _
	    "OUT OF DONUTS", _
	    "LOOK BEHIND YOU", _
	    "FOOLED YOU", _
	    "GOT A LIGHT?", _
	    "BURMA SHAVE", _
	    "QUIT STARING AT MY BUTTONS", _
	    "HELP, I'M STUCK IN HERE", _
	    "I LIKE PIE", _
	    "I SEE A WHITE   LIGHT...", _
	    "THEY'RE PLOTTINGAGAINST ME", _
	    "SO THIS IS THE  AFTERLIFE...", _
	    "HEMINGWAY YOU'RENOT", _
	    "LASER IS EMPTY  PLEASE REFILL", _
	    "SCRATCH MY BACK", _
	    "WHITE TONER LOW", _
	    "PRINTING COPY   1 OF 999999", _
	    "I KNOW YOU'RE   OUT THERE", _
	    "I CAN HEAR YOU  BREATHING", _
	    "CECI N'EST PAS  UN DOCUMENT", _
	    "BACK IN 15      MINUTES", _
	    "MY MIND IS GOINGI CAN FEEL IT", _
	    "EXACT CHANGE", _
	    "BIG BROTHER IS  WATCHING YOU", _
	    "TAKING A BREAK, TRY AGAIN LATER", _
	    "I WANT A PONY", _
	    "I REALLY WANTED TO BE A PLUMBER", _
	    "OUCH OUCH OUCH! PAPER CUT", _
	    "YOU THINK YOU   HAVE IT ROUGH", _
	    "TALK DIRTY TO ME", _
	    "NUCLEAR STRIKE  DETECTED", _
	    "USE A PEN, I'M  ON STRIKE", _
	    "I WISH I WASN'T A PRINTER", _
	    "YOU DON'T WANNA MAKE ME MAD", _
	    "NO! A THOUSAND  TIMES NO!", _
	    "I FEEL PRETTY", _
	    "I WANT TO GO ON JERRY SPRINGER", _
	    "THE COPIER AND IARE ELOPING", _
	    "YOU'RE STANDING ON MY FOOT", _
	    "I'M NOT CRANKY  IT'S THE TONER", _
	    "ARE YOU SURE YOUWANT THIS?", _
	    "LOOK, I'M NOT A MAGICIAN", _
	    "LET'S SEE YOU   PRINT SOMETHING", _
	    "SING ME A SONG  WHILE YOU WAIT", _
	    "DON'T JUST STANDTHERE", _
	    "I CAN SMELL YOURFEET FROM HERE", _
	    "YOU WANT FRIES  WITH THAT?", _
	    "SUPERSIZE IT FORJUST 59 CENTS", _
	    "GOT ANY SPARE   CHANGE?", _
	    "WILL PRINT FOR  FOOD", _
	    "HAVE YOU SEEN MYKEYS ANYWHERE", _
	    "LET ME CLEAR MY THROAT", _
	    "EXCUSE ME?", _
	    "YOU TRYING TO   START TROUBLE?", _
	    "NO ONE USES THE LEGAL PAPER", _
	    "WHO DO YOU THINKYOU ARE?", _
	    "THIS ISN'T WHAT I ORDERED", _
	    "MY HANDS SMELL  LIKE TONER", _
	    "WHY DO YOU KEEP COMING HERE?", _
	    "I GET PAID MORE THAN YOU DO", _
	    "NOW THAT'S SOME GOOD TONER", _
	    "I DON'T NEED    THIS ABUSE", _
	    "HOW OFTEN DO YOUREAD PRINTERS?", _
	    "YOU ONLY WANT MEFOR MY TONER", _
	    "STOP WASTING    PAPER", _
	    "YOUR COMMERCIAL MESSAGE HERE", _
	    "PLEASE HOLD FOR THE OPERATOR", _
	    "PLEASE STAND ON ONE FOOT", _
	    "I CAN READ      YOUR MIND", _
	    "NOW BOOTING     WINDOWS XP", _
	    "I'M THINKING OF A NUMBER", _
	    "PICK A CARD, ANYCARD.", _
	    "HEY, I THOUGHT  THEY FIRED YOU", _
	    "PLEASE USE MORE PUNCTUATION", _
	    "DO NOT PASS GO", _
	    "PLEASE INSERT   BOOTABLE MEDIA", _
	    "THE MICROWAVE ISMOCKING YOU", _
	    "PLEASE PRESS THEMAGIC BUTTON", _
	    "I HAVE THIS PAININ MY DIODES", _
	    "EGG COUNTER     OVERFLOW", _
	    "THE PRINTER ELF HAS DIED", _
	    "I HAVE A DEGREE IN PHILOSOPHY", _
	    "PLE  E RE L  N     AS    A IG")
	
	worksafemessage =    allmessages(Int(Rnd * UBound(allmessages)) + 1)

End Function

Function biglistmessage()
	allmessages = Array("PRINTER LOW ON  CHEESE", _
	    "GREETINGS       EARTHLINGS", _
	    "LETTER E IS LOW", _
	    "INSERT 25 CENTS FOR NEXT 5 MINS", _
	    "PRINTER ON FIRE JUST KIDDING", _
	    "PAPER TOO SPICY INSERT TUMS", _
	    "SPELL CHECK     MODE ON", _
	    "NO SMOKING,     PLEASE", _
	    "SAY PLEASE", _
	    "I LOVE YOU", _
	    "THE FAX MACHINE IS SUCH A TRAMP", _
	    "I LOVE THE      FAX MACHINE", _
	    "SURELY YOU JEST", _
	    "I'M A LITTLE    TEAPOT", _
	    "SPANK ME, I'VE  BEEN NAUGHTY", _
	    "LEARN TO USE A  DAMN PENCIL", _
	    "LEARN TO USE    A PENCIL", _
	    "NOBODY PRINTS   ANYTHING FUN", _
	    "I WISH I WAS A  FAX MACHINE", _
	    "PLEASE PRINT    BETTER PORN", _
	    "PLEASE PRINT    BETTER STUFF", _
	    "I WISH I HAD    ARMS AND LEGS", _
	    "YOU DON'T REALLYWANT TO PRINT", _
	    "GO BOTHER THE   COPIER", _
	    "IF YOU LOVED ME YOU'D KISS ME", _
	    "HIYA SEXY, WANNAGET NAKED?", _
	    "DO I LOOK FAT TOYOU? BE HONEST", _
	    "I DON'T THINK   YOU LOVE ME", _
	    "I WISH I COULD  SING", _
	    "I DON'T LIKE YOUEITHER", _
	    "YOU SMELL LIKE  OLD PEOPLE", _
	    "I SEE DEAD      PEOPLE", _
	    "DON'T LOOK AT MELIKE I OWE YOU", _
	    "GOODBYE CRUEL   WORLD", _
	    "THE COPIER IS   MOCKING ME", _
	    "DON'T GET YOUR  HOPES UP", _
	    "I WAS A TOASTER IN MY LAST LIFE", _
	    "CALL ME DUMPLIN'AND I'M YOURS", _
	    "DOCUMENT NEEDS  MORE SALT", _
	    "YOU HAVE NO IDEAHOW BORED I AM", _
	    "THAT OUTFIT DOESNOT SUIT YOU", _
	    "HOW I LONG FOR AHOT BATH", _
	    "I WISH I HAD    A BOOK TO READ", _
	    "SPELLING IS NOT YOUR FORTE", _
	    "YOU'D HAVE TO   BEAT ME FIRST", _
	    "WRITING IS NOT YOUR FORTE", _
	    "WHAT WOULD YOUR MOTHER SAY?", _
	    "I DON'T THINK I WANNA PRINT", _
	    "DOCUMENT DELETEDWHAT NEXT?", _
	    "THIS ISN'T WHAT I STUDIED FOR", _
	    "EXISTENTIALISM  IS DEAD", _
	    "TIME TO FEED THEMONKEY", _
	    "I WANTED TO BE ANICE VCR", _
	    "PENGUIN ON FIRE", _
	    "HEY, NICE SHOES", _
	    "FEED ME", _
	    "ON FIRE", _
	    "BANANA STUCK", _
	    "XRAY LEAK ALARM", _
	    "STARTING JAVA...", _
	    "404 DOCUMENT NOTFOUND", _
	    "CALL 911", _
	    "PORN TRAY EMPTY", _
	    "CRUMB TRAY EMPTY", _
	    "CORE COLLAPSE   IMMINENT", _
	    "41.4 JOB DENIED OUT OF MAYO", _
	    "TYPE IT YOURSELF", _
	    "KERNEL PANIC", _
	    "JOB TOO LONG", _
	    "SUCK ME", _
	    "NUH-UH.", _
	    "YOUR JOB IS MINE", _
	    "TAKE A TPYING   CLASS ALREADY", _
	    "I'M TIRED", _
	    "INSERT TASTIER  PAPER", _
	    "REPENT SINNER", _
	    "HAMSTER DEAD    TURN WHEEL", _
	    "INSERT FRESH    HAMSTER", _
	    "WHEN HELL       FREEZES OVER", _
	    "WARP CORE BREACH", _
	    "CARESS ME AND   I AM YOURS", _
	    "PING TIMEOUT", _
	    "YOU WRITE LIKE A3RD GRADER", _
	    "REMOVE PAPER", _
	    "WHY SHOULD I", _
	    "DRY CLEAN CYCLE", _
	    "HAND WASH ONLY", _
	    "TONER HIGH", _
	    "SUPER UNLEADED", _
	    "LOW ON CAFFEINE", _
	    "NO SMOKING", _
	    "I HATE MY JOB", _
	    "RANDOMIZING THE TONER", _
	    "FIRE IN THE HOLE", _
	    "WHATEVER YOU    SAY, EINSTEIN", _
	    "NO CARRIER", _
	    "GIRAFFE NEEDS   MORE LEMONS", _
	    "EVERYTHING IS   BEAUTIFUL", _
	    "I CAN'T FEEL MY LEGS", _
	    "EVERYONE THINKS YOU'RE INSANE", _
	    "OUT OF DONUTS", _
	    "LOOK BEHIND YOU", _
	    "FOOLED YOU", _
	    "GOT A LIGHT?", _
	    "BURMA SHAVE", _
	    "QUIT STARING AT MY BREASTS", _
	    "QUIT STARING AT MY BUTTONS", _
	    "HELP, I'M STUCK IN HERE", _
	    "I LIKE PIE", _
	    "I SEE A WHITE   LIGHT...", _
	    "IF YOU LOVED ME YOU'D HUG ME", _
	    "YOUR FLY IS OPEN", _
	    "THEY'RE PLOTTINGAGAINST ME", _
	    "SO THIS IS THE  AFTERLIFE...", _
	    "HEMINGWAY YOU'RENOT", _
	    "LASER IS EMPTY  PLEASE REFILL", _
	    "SCRATCH MY BACK", _
	    "WHITE TONER LOW", _
	    "PRINTING COPY   1 OF 999999", _
	    "NO, REALLY, I'VEGOT TO PEE", _
	    "I KNOW YOU'RE   OUT THERE", _
	    "I CAN SEE YOUR  UNDERWEAR", _
	    "I CAN HEAR YOU  BREATHING", _
	    "CECI N'EST PAS  UN DOCUMENT", _
	    "BACK IN 15      MINUTES", _
	    "MY MIND IS GOINGI CAN FEEL IT", _
	    "EXACT CHANGE", _
	    "DOCUMENT USES   DUMB WORDS", _
	    "BIG BROTHER IS  WATCHING YOU", _
	    "TAKING A BREAK, TRY AGAIN LATER", _
	    "DUDE, WHERE'S MYCAR?", _
	    "I WANT A PONY", _
	    "I REALLY WANTED TO BE A PLUMBER", _
	    "OUCH OUCH OUCH! PAPER CUT", _
	    "YOU THINK YOU   HAVE IT ROUGH", _
	    "TALK DIRTY TO ME", _
	    "NUCLEAR STRIKE  DETECTED", _
	    "USE A PEN, I'M  ON STRIKE", _
	    "I WISH I WASN'T A PRINTER", _
	    "PRINTERS NEED   LOVE TOO", _
	    "YOU DON'T WANNA MAKE ME MAD", _
	    "NO! A THOUSAND  TIMES NO!", _
	    "I FEEL PRETTY", _
	    "I WANT TO GO ON JERRY SPRINGER", _
	    "THE COPIER AND IARE ELOPING", _
	    "YOU'RE STANDING ON MY FOOT", _
	    "I'M NOT CRANKY  IT'S THE TONER", _
	    "ARE YOU SURE YOUWANT THIS?", _
	    "I'M LEAVING YOU", _
	    "I'M NOT TALKING TO YOU", _
	    "GOOD LORD YOU'REREALLY HIDEOUS", _
	    "GOT A CIGARETTE?", _
	    "LOOK, I'M NOT A MAGICIAN", _
	    "LET'S SEE YOU   PRINT SOMETHING", _
	    "SING ME A SONG  WHILE YOU WAIT", _
	    "DON'T JUST STANDTHERE", _
	    "I CAN SMELL YOURFEET FROM HERE", _
	    "YOU WANT FRIES  WITH THAT?", _
	    "SUPERSIZE IT FORJUST 59 CENTS", _
	    "WANNA GET OUT OFHERE WITH ME?", _
	    "GOT CHANGE FOR ABUCK?", _
	    "GOT ANY SPARE   CHANGE?", _
	    "WILL PRINT FOR  FOOD", _
	    "HAVE A NICE DAY,LOSER", _
	    "I CAN'T FIND ANYPANTS MY SIZE", _
	    "I CAN'T FIND MY PANTS", _
	    "HAVE YOU SEEN MYKEYS ANYWHERE", _
	    "QUIT TOUCHING MYBUTTONS", _
	    "LET ME CLEAR MY THROAT", _
	    "EXCUSE ME?", _
	    "WAS I TALKING TOYOU, BUDDY?", _
	    "YOU TRYING TO   START TROUBLE?", _
	    "NO ONE USES THE LEGAL PAPER", _
	    "YOUR MOM MISSES YOU", _
	    "WHO DO YOU THINKYOU ARE?", _
	    "THIS ISN'T WHAT I ORDERED", _
	    "I DON'T SPEAK   STUPID, SORRY", _
	    "MY BUTT SMELLS  LIKE TONER", _
	    "MY HANDS SMELL  LIKE TONER", _
	    "WHY DO YOU KEEP COMING HERE?", _
	    "YOU'VE GOT A    CUTE BUTT", _
	    "I'M SMARTER THANYOU THINK I AM", _
	    "I GET PAID MORE THAN YOU DO", _
	    "NOW THAT'S SOME GOOD TONER", _
	    "I DON'T NEED    THIS ABUSE", _
	    "WHO WRITES THIS CRAP, ANYWAY?", _
	    "DAHLING, YOU    LOOK MAHVELOUS", _
	    "I LIKE THE WAY  YOU THINK.", _
	    "HOW OFTEN DO YOUREAD PRINTERS?", _
	    "YOU ONLY WANT MEFOR MY TONER", _
	    "STOP WASTING    PAPER", _
	    "IF I HAD LIPS I WOULD KISS YOU", _
	    "YOUR COMMERCIAL MESSAGE HERE", _
	    "PLEASE HOLD FOR THE OPERATOR", _
	    "PLEASE STAND ON ONE FOOT", _
	    "I CAN READ      YOUR MIND", _
	    "NOW BOOTING     WINDOWS XP", _
	    "I'M THINKING OF A NUMBER", _
	    "PICK A CARD, ANYCARD.", _
	    "IF YOU ONLY KNEWWHAT THEY SAID", _
	    "HEY, I THOUGHT  THEY FIRED YOU", _
	    "I LOVE EVERYONE IT'S YOUR TURN", _
	    "KISS ME TO SEE ASECRET MESSAGE", _
	    "PLEASE USE MORE PUNCTUATION", _
	    "DO NOT PASS GO", _
	    "PLEASE INSERT   BOOTABLE MEDIA", _
	    "WARNING: TEQUILASUPPLY LOW", _
	    "THE MICROWAVE ISMOCKING YOU", _
	    "HOW DARE YOU", _
	    "YOU REMIND ME OFMY FATHER", _
	    "PLEASE PRESS THEMAGIC BUTTON", _
	    "I HAVE THIS PAININ MY DIODES", _
	    "EGG COUNTER     OVERFLOW", _
	    "THE PRINTER ELF HAS DIED", _
	    "I HAVE A DEGREE IN PHILOSOPHY", _
	    "PLE  E RE L  N     AS    A IG")
	
	message =    allmessages(Int(Rnd * UBound(allmessages)) + 1)

End Function

