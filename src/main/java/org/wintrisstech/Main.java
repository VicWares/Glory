package org.wintrisstech;
/****************************************
 * Glory...new start combind Crazy2 with NewCovers...both work sort of
 * version Glory 220818A
 ****************************************/
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.swing.*;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
public class Main
{
    private static final String VERSION = "220818A";
    private XSSFWorkbook sportDataWorkbook;
    private HashMap<String, String> weekDateMap = new HashMap<>();
    private HashMap<String, String> weekDateMap2 = new HashMap<>();
    private HashMap<String, String> cityNameMap = new HashMap<>();
    private HashMap<String, String> xRefMap = new HashMap<>();
    private WebSiteReader webSiteReader = new WebSiteReader();
    public ExcelReader excelReader = new ExcelReader();
    public ExcelBuilder excelBuilder = new ExcelBuilder();
    public ExcelWriter excelWriter = new ExcelWriter();
    public DataCollector dataCollector = new DataCollector();
    public WebSiteReader websiteReader;
    private org.jsoup.select.Elements consensusElements;
    private int excelRowIndex = 3;
    private String dataGame;
    private Elements bet365Elements;
    private String homeTeamShortName;
    private String awayTeamShortName;
    public static void main(String[] args) throws IOException, InterruptedException
    {
        System.out.println("Version Glory " + VERSION);
        Main main = new Main();
        main.getGoing();//To get out of static context
    }
    private void getGoing() throws IOException, InterruptedException
    {
        websiteReader = new WebSiteReader();
        {
            fillCityNameMap();//2022-2023 season
            excelBuilder.setSeason("2022");
        }
//        {
//        fillWeekDateMap2();//2021-2022 season
//        excelBuilder.setSeason("2021");
//        }
        excelBuilder.setCityNameMap(cityNameMap);
        String weekNumber = JOptionPane.showInputDialog("Enter NFL week number");
        weekNumber = "1";//For testing
        excelBuilder.setWeekNumber(weekNumber);
        String weekDate = weekDateMap.get(weekNumber);//Gets week date e.g. 2022-09-08 from week number e.g. 1,2,3,4,...
        org.jsoup.select.Elements nflElements = webSiteReader.readWebsite("https://www.covers.com/sports/nfl/matchups?selectedDate=" + weekDate);//Covers.com "Scores and Matchups" page for this week
        org.jsoup.select.Elements soupOddsElements = webSiteReader.readWebsite("https://www.covers.com/sport/football/nfl/odds");
        org.jsoup.select.Elements weekElements = nflElements.select(".cmg_game_data, .cmg_matchup_game_box");//Lots of good stuff in this Element: team name, team city...
        xRefMap = buildXref(weekElements);//Key is data-event-ID e.g 87579, Value is data-game e.g 265282, two different ways of selecting the same matchup (game)
        System.out.println(xRefMap);
        System.out.println("Main57 week number => " + weekNumber + ", week date => " + weekDate + ", " + weekElements.size() + " games this week");
        sportDataWorkbook = excelReader.readSportData();
        System.out.println("Main52 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> BEGIN MAIN LOOP  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ");
        for (Map.Entry<String, String> entry : xRefMap.entrySet())
        {
            System.out.println("-------------------------------------------------------------------------------------------------------------------------------------------------------------------BEGIN");
            String dataEventId = entry.getKey();
            dataGame = xRefMap.get(dataEventId);
            awayTeamShortName = nflElements.select("[data-event-id='" + dataEventId + "']").attr("data-away-team-shortname-search");
            homeTeamShortName = nflElements.select("[data-event-id='" + dataEventId + "']").attr("data-home-team-shortname-search");
            System.out.println("Main67, data-event-id=> " + dataEventId + ", data-game=> " + dataGame + ", " + awayTeamShortName + "/" + homeTeamShortName);
            consensusElements = webSiteReader.readWebsite("https://contests.covers.com/consensus/matchupconsensusdetails/dc2b41af-f52f-4e17-b0b1-ac2900676797?showExperts=" + dataEventId);
            dataCollector.collectConsensusData(consensusElements, dataEventId);
            excelBuilder.setConsensusElements(consensusElements);
//            excelBuilder.setGameDatesMap(dataCollector.getGameDatesMap());
//            excelBuilder.setAtsHomesMap(dataCollector.getAtsHomesMap());
//            excelBuilder.setAtsAwaysMap(dataCollector.getAtsAwaysMap());
//            excelBuilder.setOuOversMap(dataCollector.getOuOversMap());
//            excelBuilder.setOuUndersMap(dataCollector.getOuUndersMap());
            excelBuilder.setGameIdentifier(dataCollector.getGameIdentifierMap().get(dataEventId));
            excelBuilder.buildExcel(sportDataWorkbook, dataEventId, dataGame, excelRowIndex, soupOddsElements, nflElements);
            System.out.println("Main75=====> dataEventId " + dataEventId + " dataGame, " + dataGame + "  is " + awayTeamShortName + "/" + homeTeamShortName + "<==== dataGame " + dataGame);
            excelRowIndex++;
            System.out.println("------------------------------------------------------------------------------------------------------------------------------------------------------------------END");
            System.out.println("777777");
        }
        System.out.println("Main70 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> END MAIN LOOP  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ");
        excelWriter.openOutputStream();
        excelWriter.writeSportData(sportDataWorkbook);
        excelWriter.closeOutputStream();
        System.out.println("Main82 Proper Finish...HOORAY!");
    }

    public HashMap<String, String> buildXref(org.jsoup.select.Elements weekElements)
    {
        for (Element e : weekElements)
        {
            String dataLinkString = e.attr("data-link");
            String[] dlsa = dataLinkString.split("/");
            String dataLink = dlsa[5];
            String dataEvent = e.attr("data-event-id");
            xRefMap.put(dataEvent, dataLink);
        }
        return xRefMap;
    }
    private void fillCityNameMap()
    {
        cityNameMap.put("Minneapolis", "Minnesota");//Minnesota Vikings
        cityNameMap.put("Tampa", "Tampa Bay");//Tampa Bay Buccaneers
        cityNameMap.put("Tampa Bay", "Tampa Bay");//Tampa Bay Buccaneers
        cityNameMap.put("Arlington", "Dallas");//Dallas Cowboys
        cityNameMap.put("Dallas", "Dallas");//Dallas Cowboys
        cityNameMap.put("Orchard Park", "Buffalo");//Buffalo Bills
        cityNameMap.put("Buffalo", "Buffalo");//Buffalo Bills
        cityNameMap.put("Charlotte", "Carolina");//Carolina Panthers
        cityNameMap.put("Carolina", "Carolina");//Carolina Panthers
        cityNameMap.put("Arizona", "Arizona");//Arizona Cardinals
        cityNameMap.put("Tempe", "Arizona");//Arizona Cardinals
        cityNameMap.put("Foxborough", "New England");//New England Patriots
        cityNameMap.put("New England", "New England");//New England Patriots
        cityNameMap.put("East Rutherford", "New York");//New York Giants and New York Jets
        cityNameMap.put("New York", "New York");//New York Giants and New York Jets
        cityNameMap.put("Landover", "Washington");//Washington Football Team
        cityNameMap.put("Washington", "Washington");//Washington Football Team
        cityNameMap.put("Nashville", "Tennessee");//Tennessee Titans
        cityNameMap.put("Miami", "Miami");//Miami Dolphins
        cityNameMap.put("Baltimore", "Baltimore");//Baltimore Ravens
        cityNameMap.put("Cincinnati", "Cincinnati");//Cincinnati Bengals
        cityNameMap.put("Cleveland", "Cleveland");//Cleveland Browns
        cityNameMap.put("Pittsburgh", "Pittsburgh");//Pittsburgh Steelers
        cityNameMap.put("Houston", "Houston");//Houston Texans
        cityNameMap.put("Indianapolis", "Indianapolis");//Indianapolis Colts
        cityNameMap.put("Jacksonville", "Jacksonville");//Jacksonville Jaguars
        cityNameMap.put("Tennessee", "Tennessee");//Tennessee Titans
        cityNameMap.put("Denver", "Denver");//Denver Broncos
        cityNameMap.put("Kansas City", "Kansas City");//Kansas City Chiefs
        cityNameMap.put("Las Vegas", "Las Vegas");//Los Angeles Chargers and Los angeles Rams
        cityNameMap.put("Philadelphia", "Philadelphia");//Philadelphia Eagles
        cityNameMap.put("Chicago", "Chicago");//Chicago Bears
        cityNameMap.put("Detroit", "Detroit");//Detroit Lions
        cityNameMap.put("Green Bay", "Green Bay");//Green Bay Packers
        cityNameMap.put("Minnesota", "Minnesota");
        cityNameMap.put("Atlanta", "Atlanta");//Atlanta Falcons
        cityNameMap.put("New Orleans", "New Orleans");//New Orleans Saints
        cityNameMap.put("Los Angeles", "Los Angeles");//Los Angeles Rams
        cityNameMap.put("San Francisco", "San Francisco");//San Francisco 49ers
        cityNameMap.put("Seattle", "Seattle");//Seattle Seahawks
    }
    private void fillWeekDateMap()
    {
        weekDateMap.put("1", "2022-09-08");//Season 2022 start...Week 1
        weekDateMap.put("2", "2022-09-15");
        weekDateMap.put("3", "2022-09-22");
        weekDateMap.put("4", "2022-09-29");
        weekDateMap.put("5", "2022-10-06");
        weekDateMap.put("6", "2022-10-13");
        weekDateMap.put("7", "2022-10-20");
        weekDateMap.put("8", "2022-10-27");
        weekDateMap.put("9", "2022-11-03");
        weekDateMap.put("10", "2022-11-10");
        weekDateMap.put("11", "2022-11-17");
        weekDateMap.put("12", "2022-11-24");
        weekDateMap.put("13", "2022-12-01");
        weekDateMap.put("14", "2022-12-08");
        weekDateMap.put("15", "2022-12-15");
        weekDateMap.put("16", "2022-12-22");
        weekDateMap.put("17", "2022-12-29");
        weekDateMap.put("18", "2023-01-08");
        weekDateMap.put("19", "2023-02-05");
    }
    private void fillWeekDateMap2()//2021 season for testing
    {
        weekDateMap.put("1", "2021-09-09");//Season 2021 start...Week 1
        weekDateMap.put("2", "2022-09-16");
        weekDateMap.put("3", "2021-09-23");
        weekDateMap.put("4", "2021-09-30");
        weekDateMap.put("5", "2021-10-07");
        weekDateMap.put("6", "2021-10-14");
        weekDateMap.put("7", "2021-10-21");
        weekDateMap.put("8", "2021-10-28");
        weekDateMap.put("9", "2021-11-04");
        weekDateMap.put("10", "2021-11-11");
        weekDateMap.put("11", "2021-11-18");
        weekDateMap.put("12", "2021-11-25");
        weekDateMap.put("13", "2021-12-02");
        weekDateMap.put("14", "2021-12-09");
        weekDateMap.put("15", "2021-12-16");
        weekDateMap.put("16", "2021-12-23");
        weekDateMap.put("17", "2021-01-02");
        weekDateMap.put("18", "2022-01-08");
    }
}
