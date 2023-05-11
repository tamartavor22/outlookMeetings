Param(
    [Parameter(Mandatory, HelpMessage = "Please input a start date like this: dd/mm/yyyy")] [string]$StartDate,
    [Parameter(Mandatory, HelpMessage = "Please input an end date like this: dd/mm/yyyy")] [string]$EndDate
)

Function Get-DiunimDate ($mapi)
# Gets dates of diunim days
{
#the below code doesnt give me recurring events. only the first of them...
$DiunimFilter = "[MessageClass]='IPM.Appointment' AND [Subject] = 'מוקד המשכים' AND [AllDayEvent] = 'true'"
$Appointments = ($mapi.GetDefaultFolder(9).Items).restrict($DiunimFilter)
$Appointments = $Appointments | select start
#the below code works good as a filter
$Appointments.start -gt $StartDate -lt $EndDate
}

Function Get-Events($relevantDate, $mapi)
# finds all events on a diun day
{
    $diunimDate = $relevantDate.AddDays(-1).toString('d/M/yyyy')
    $diunimendate = $relevantDate.AddDays(1).toString('d/M/yyyy')

    $eventsFilter = "[MessageClass]='IPM.Appointment' AND [Start] > '$diunimDate'  AND [Start] < '$diunimendate' AND [isrecurring] = 'False' AND [AllDayEvent] = 'false'"
    $events = ($mapi.GetDefaultFolder(9).Items).restrict($eventsFilter)

    $events | select start, end, subject
    Write-Host $events
   
}

Function Get-WorkingHoursInDate ($relevantDate)
# Creates an array with datetime object of each half hour in relevant workday times.
# Ommiting hour of break
{
    $workdayStart = Get-Date -Date "$($relevantDate.ToShortDateString()) 08:30:00"
    $workdayEnd = Get-Date -Date "$($relevantDate.ToShortDateString()) 14:30:00"

    $workHour = $workdayStart
    $allWorkHours = @($workdayStart)

    while ($workHour -le $workdayEnd)
        {$workHour = $workHour.AddMinutes(30)
        if (-Not (($workHour.hour -eq 11) -and ($workHour.Minute -eq 0))) {
            $allWorkHours += $workHour
        }
    }
    $allWorkHours

}

Function Get-FreeHours ($eventsInRelevantDate, $relevantDate, $allWorkHours)
# For each working hour looks to see how many events are already booked.
# If less then 4, adds it to free time array
{
    $groupedEventsInHour = $eventsInRelevantDate.start | group
    $freeTime = @()

    Foreach ($workHour in $allWorkHours) {
        $amountOfEventsInHour = $($groupedEventsInHour | Where-Object {$_.Name -eq $workHour.toString()}).count
        if ($amountOfEventsInHour -lt 4) {
            $freeTime += $workHour
        }
    }
   
    $freeTime
}


Function Get-AllFreeTime
# Runs main logic for each day of diunim found and gathers to one array all available dates and timess
{
    $outlook=New-Object -com outlook.application
    $mapi=$outlook.GetNamespace("MAPI")
    $relevantDates = Get-DiunimDate($mapi)
    Write-Host $relevantDates

    if ($relevantDates.length -eq 0) {
        Write-Host "No Diunim Dates In Selected Interval... :( Love You Mom"
    }
    else {
        $allFreeTime = @()

        Foreach ($relevantDate in $relevantDates) {
            $eventsInRelevantDate = Get-Events -relevantDate $relevantDate -mapi $mapi
            $allWorkHours = Get-WorkingHoursInDate($relevantDate)
            $freeTime = Get-FreeHours -eventsInRelevantDate $eventsInRelevantDate -relevantDate $relevantDate -allWorkHours $allWorkHours
            $allFreeTime += $freeTime
        }
        $allFreeTime
    }
}

$my_html1 = @'
<!DOCTYPE html>
<html>
  <head>
  <meta charset="UTF-8"/>
    <style type='text/css' media='screen'>
      /* yellowgreen, orange */
      body {
        /* background: beige; */
        font-family: 'system-ui';
      }

      table,
      th,
      td {
        border: 1px solid;
        border-collapse: collapse;
        border-color: lightgrey;
      }

      ul {
        list-style-type: none;
        padding-inline-start: 0px;
      }

      li.hours {
        display: inline-block;
        padding-left: 5px;
        padding-right: 5px;
        margin-left: 5px;
        margin-bottom: 5px;
        border-radius: 5px;
        color: white;
      }

      li.hours:hover {
        cursor: pointer;
        color: black;
        background-color: white !important;
      }

      td {
        font-size: large;
        text-align: center;
        vertical-align: middle;
        min-width: 80px;
        max-width: 300px;
        max-height: 100%;
        text-overflow: ellipsis;
      }

      table {
        direction: rtl;
        /* background: whitesmoke; */
      }
    </style>
  </head>
  <body>
    <center>
      <h1>דיוני מוקד</h1>
      <h2 style='margin-bottom: 40px'>כבוד השופטת נעה תבור</h2>
      <div id='myChosenDate'></div>
      <div id='myTable'></div>
    </center>

    <script>
      let tableElement = document.getElementById('myTable');
      let tableContent =
        '<table><tr><th>ראשון</th><th>שני</th><th>שלישי</th><th>רביעי</th><th>חמישי</th></tr><tr>';

      let dataStr =
'@

$my_html2 = @'
;

      dataArray = dataStr.map(
        (d) => new Date(Number(d.value.replace("/Date(", "").replace(")/", "")))
      );
            dataArray.sort((date1, date2) => date1 - date2);


      function dateToString(date) {
        const dateStr = `${date.getDate()}/${date.getMonth()}/${date.getYear()}`;
        const timeStr = `${date.getHours()}:${String(
          date.getMinutes()
        ).padStart(2, "0")}`;
        return [dateStr, timeStr];
      }

      //   Filling first row of table with empty cells up to the first event day
      let currentEventDate = dataArray[0];
      let [currentEventDateStr, currentEventTimeStr] =
        dateToString(currentEventDate);
      let currentEventDayOfWeek = currentEventDate.getDay();
      tableContent += "<td>חסום</td>".repeat(currentEventDayOfWeek);
      tableContent += `<td><ul id=${currentEventDateStr}><li style='margin-bottom: 8px'>${String(
        currentEventDateStr
      )}</li><li class='hours'>${currentEventTimeStr}</li>`;

      // looping through each meeting event
      for (var i = 1; i < dataArray.length; i++) {
        let newEventDate = dataArray[i];
        let newEventDayOfWeek = newEventDate.getDay();
        let daysBetweenEvents = (newEventDate - currentEventDate) / 86400000; //number of miliseconds in a day
        let [newEventDateStr, newEventTimeStr] = dateToString(newEventDate);
        let daysToEndOfCurrentWeek = 6 - currentEventDayOfWeek;

        // if 2 meeting are in the same day:
        if (daysBetweenEvents < 1) {
          //   rounding down all events in same hour to one event
          tableContent += `<li class='hours'>${newEventTimeStr}</li>`;
          currentEventDate = newEventDate;
        }

        //  if between 2 meetings passed less than a week and not in the same day:
        if (
          currentEventDayOfWeek + daysBetweenEvents < 6 &&
          daysBetweenEvents > 1
        ) {
          console.log(newEventDayOfWeek, currentEventDayOfWeek);
          tableContent += "</ul></td>";
          tableContent += "<td>חסום</td>".repeat(
            newEventDayOfWeek - currentEventDayOfWeek - 1
          );
          tableContent += `<td><ul id=${newEventDateStr}><li style='margin-bottom: 8px'>${newEventDateStr}</li><li class='hours'>${newEventTimeStr}</li>`;
          currentEventDate = newEventDate;
          currentEventDayOfWeek = newEventDayOfWeek;
        }

        // if between 2 following meetings the week ended:
        if (currentEventDayOfWeek + daysBetweenEvents > 6) {
          tableContent += "</ul></td>";

          //   closing the previos events week
          tableContent += "<td>חסום</td>".repeat(daysToEndOfCurrentWeek - 2); //minus 2 beacuse i am not showing days friday and saturday
          tableContent += "</tr><tr>";

          // if between 2 following meetings passed more than 1 week:
          if (daysBetweenEvents - daysToEndOfCurrentWeek > 7) {
            // filling blocked days for each full empty week between the events
            for (
              var w = 0;
              w < (daysBetweenEvents - daysToEndOfCurrentWeek) / 7 - 1;
              w++
            ) {
              tableContent += "<td>חסום</td>".repeat(5);
              tableContent += "</tr><tr>";
            }
          }

          //   adding the current event while creating the new week and filling it with the appropriate ammount of days passed from sunday
          tableContent += "<td>חסום</td>".repeat(newEventDayOfWeek);
          tableContent += `<td><ul id=${newEventDateStr}><li style='margin-bottom: 8px'>${newEventDateStr}</li><li class='hours'>${newEventTimeStr}</li>`;

          currentEventDate = newEventDate;
          currentEventDayOfWeek = newEventDayOfWeek;
        }
      }
      tableContent += "</tr></table>";
      tableElement.innerHTML = tableContent;

      //   Handeling choosing a date
      let chosenDateElement = document.getElementById("myChosenDate");

      function createElementChosenDate(chosenHour, chosenDate) {
        let chosenDateData = `<h2 style='background-color: #e67c73; color: white; padding-top: 10px; padding-bottom:10px; margin-bottom: 35px; border-radius: 5px'>הדיון הבא נקבע לתאריך ${chosenDate} בשעה ${chosenHour}</h2>`;
        chosenDateElement.innerHTML = chosenDateData;
      }

      const hourElements = document.getElementsByClassName("hours");
      for (hourElement = 0; hourElement < hourElements.length; hourElement++) {
        hourElements[hourElement].onclick = function (e) {
          createElementChosenDate(
            e.target.innerHTML,
            e.target.parentElement.id
          );
        };
      }

      //   coloring each days hours in the same random color

      let datesUls = document.getElementsByTagName("ul");
      const colorList = [
        "#f6bf26",
        "#3f51b5",
        "#33b679",
        "#8e24aa",
        "#039be5",
        "#7986cb",
        "#d50000",
      ];

      for (ul = 0; ul < datesUls.length; ul++) {
        for (li = 1; li < datesUls[ul].children.length; li++) {
          datesUls[ul].children[li].style.backgroundColor =
            colorList[ul % colorList.length];
        }
      }
    </script>
  </body>
</html>

'@


$allFreeTime = Get-AllFreeTime
Write-Host $allFreeTime
$allFreeTimeStr = $($allFreeTime | ConvertTo-Json)
$full_html = $my_html1 + $allFreeTimeStr + $my_html2
#$full_html > "C:\Users\noata\Downloads\OutlookEventsTest.html"
Out-File -FilePath "C:\Users\noata\Downloads\OutlookEventsTest.html" -Force -InputObject $full_html -Encoding UTF8
Write-Host "Finished running :)"
& 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe' 'C:\Users\noata\Downloads\OutlookEventsTest.html'
