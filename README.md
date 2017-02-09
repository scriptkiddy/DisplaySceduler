# DisplaySceduler

Wenn man die DisplaySceduler.exe ausführt, wird ein gegebenes Verzeichnis nach Powerpointdateien durchsucht. Aus diesen Dateien wird nach einer Namenskonven-
tion eine Datei ausgesucht und diese Datei wird dann im Vollbild auf dem Display angezeigt.
Bevor die ausgewählte Datei angezeigt wird, werden zwei Änderungen an der Datei vorgenommen:
1. Der Powerpointfoliensatz wird so eingestellt, dass er wieder von vorne beginnt, wenn er am Ende angekommen ist.
2. Bei allen Folien in denen kein automatischer Übergang nach einer bestimmten Zeit eingestellt ist, wird ein Automatischer Übergang nach einer bestimmten Anzahl an Sekunden eingestellt. (Die Sekundenanzahl ist per parameter einstellbar)
Diese beiden Änderungen sollen häufige Fehler, wegen denen der Foliensatz auf einer Folie "hängen bleibt" verhindern.

## Hinweise für Nutzer und Administratoren:

### Namenskonvention der Dateien:
Alle Dateinamen beginnen mit dem Wort "play", dahinter folgt eine Datumsangabe und als Dateiendung wird wie gewohnt "ppt" oder "pptx" verwendet.

#### 1.1 einzelner Tag
	Wenn der Foliensatz an einem bestimmten ganzen Tag gezeigt werden soll, kann das Datum einfach im amerikanischen Datumsformat angehängt werden.

	Konvention:	playYYYYMMDD.ppt(x)
	Beispiel:	  play20170208.ppt
	
	Diese Datei soll am 8. Februar 2017 abgespielt werden.

#### 1.2 bestimmter Zeitraum
	Wenn der Foliensatz einen bestimmten Zeitraum lang angezeigt werden soll, kann auch der Zeitraum minutengenau angegeben werden.
	
	Konvention: playYYYYMMDD_HHMM-YYYYMMDD_HHMM.ppt(x)
	Beispiel: play20170101_0000-20171231_2359.ppt
	
	Diese Datei wird das ganze Jahr 2017 angezeigt. 
	
	Beispiel: play20170208_1530-20170208_2000.ppt
	
	Diese Datei wird am 8.  Februar von 15:30 bis 20:00 Uhr angezeigt.
_!Achtung es müssen immer auch Minuten angegeben werden!_

Wenn mehrere Dateien zum aktuellen Zeitpunkt passen wird immer die Datei mit dem __kleinsten Zeitraum__ ausgewählt. Am 8. Februar 2017 um 15:45 Uhr würde also aus den Dateien oben die letzte Datei ausgewählt werden weil ihr Zeitrum nur wenige Stunden umfasst.

### Hinweis:
es hat sich als sinnvolle Praxis ergeben, dass man einen Foliensatz erstellt der das ganze Jahr angezeigt werden soll. In diesen Foliensatz tut man einfach all die Folien die jederzeit in Rotation angezeigt werden sollen. 
Diesen Foliensatz kann man immer wieder verändern, so dass er zum Beispiel die aktuellen Ankündigungen und allgemeine Hinweise enthält. Wenn man nun zusätzlich zu dieser allgmeinen Rotation für einzelne Ereignisse Ausnahmen festlegen will kann man einfach für bestimmte Tage oder Zeiträume weitere Foliensätze in den gleichen Ordner ablegen.

## Hinweise für Administratoren:
Die aktuellen Quellen für den DisplaySceduler finden Sie immer auf: [GitHub](https://github.com/scriptkiddy/DisplaySceduler.git)
### Kommandozeilenparameter:
<table>
	<tr>
		<td>-v</td>
		<td>Debugging einschalten, dieser Parameter sollte nicht im Betrieb verwendet werden</td>
	<tr>
		<td>-d _Zahl_</td>
		<td>Mit dieser Zahl kann eine Azahl von Sekunden angegeben werden die als Zeitdauer für Folien ohne voreingestellte Zeitdauer  verwendet werden soll</td>
	<tr>
		<td width="200">-p _Ordnerpfad_</td>
		<td>Ordnerpfad zum dem Ordner in dem die Powerpointdatien liegen die angezeigt werden sollen</td>
	</tr>
</table>
### default.ppt
Im gleichen Ordner wie das Programm muss noch eine default.ppt Datei Hinterlegt sein. Diese wird angezeigt, wenn nichts anderes angezeigt werden kann. Dort Sollte also ein allgemeiner Hintergrund zu sehen sein oder vielleicht eine Kontaktinformation von demjenigen der das Display verwaltet.

### Best Practice:
DisplaySceduler muss auf einem Computer mit Windows 7 oder neuer verwendet werden. Außerdem muss Powerpoint installiert sein. Sinnvollerweise startet man den
Displaysceduler im Windows Autostart per Script (.bat-Datei) und stellt im  Windows für den Benutzer Autloogin ein. Den Ordner mit den Foliensätzen holt man entweder aus einem Netzlaufwerk oder gibt den Ordner per Windowsfreigabe
frei. Dann braucht der Redakteur des Displays später nur Dateien in dem Ordner bearbeiten. Der angezeigte Foliensatz wird immer nur beim Start des DisplayScedulers ausgewählt. Wenn man möchte, dass das Auch tagsüber die Angezeigt 
Folie wechseln kann, sollte man einen Windows Task erstellen der alle x Minuten läuft. Dann wird alle x Minuten erneut der Foliensatz ausgewählt. Wenn man mit den Beispielen oben den Displaysceduler alle 20 Minuten (00, 20 40) starten  würde, würde also um 15:40 Uhr am 8. Februar 2017 die Folie für 15:30 bis 20:00 ausgewählt werden davor die Folie für den ganzen Tag.

### Ordnerinhalt
<table>
	<tr>
		<td> DisplaySceduler </td>
		<td>:</td>
		<td>Ordner mit dem VisualStudio Projekt</td>
	</tr>
	<tr>
		<td>slides</td>
		<td>:</td>
		<td>Ordner mit Beispiel Foliensätzen (inkl. default.ppt)</td>
	</tr>
	<tr>
		<td>DisplaySceduler.sln</td>
		<td>:</td>
		<td>VisualStudio Projekt-Datei</td>
	</tr>
	<tr>
		<td>README.md</td>
		<td>:</td>
		<td>die README-Datei</td>
	</tr>
</table>
