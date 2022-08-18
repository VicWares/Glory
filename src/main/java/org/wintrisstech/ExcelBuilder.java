package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220817
 *******************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.select.Elements;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
public class ExcelBuilder
{
    private String season;
    private String weekNumber;
    private String ouUnder;
    private String ouOver;
    private String homeCity;
    private String awayTeam;
    private String thisMatchupDate;
    private String atsHome;
    private String atsAway;
    private String completeHomeTeamName;
    private String completeAwayTeamName;
    private String gameIdentifier;
    private String awayMoneyLineOdds;
    private String homeMoneyLineOdds;
    private String awaySpreadCloseOdds;
    private String homeSpreadOdds;
    private String awayNickname;
    private String homeNickname;
    private HashMap<String, String> gameDatesMap = new HashMap<>();
    private HashMap<String, String> atsHomesMap = new HashMap<>();
    private HashMap<String, String> atsAwaysMap = new HashMap<>();
    private HashMap<String, String> ouOversMap;
    private HashMap<String, String> ouUndersMap;
    private HashMap<String, String> homeMLOddsMap = new HashMap<>();
    private Sheet sportDataSheet;
    private XSSFWorkbook sportDataWorkBook = new XSSFWorkbook();
    private XSSFSheet sportDataUpdateSheet = null;
    private HashMap<String, String> homeMoneyLineOddsMap = new HashMap<>();
    private HashMap<String, String> awayMoneyLineOddsMap = new HashMap<>();
    private HashMap<String, String> homeSpreadOddsMap = new HashMap<>();
    private HashMap<String, String> awaySpreadOddsMap = new HashMap<>();
    private String awayCity;
    private String homeTeamShortName;
    private String awayTeamShortName;
    private String awayTeamSpreadOpenOdds;
    private String homeMoneylineCloseOdds;
    private String homeSpreadOpenOdds;
    private String homeSpreadCloseOdds;
    private String awayMoneylineCloseOdds;
    private Elements xx;
    private Elements dataEventIdElements;
    private Elements bet365DataGameElements;
    private String awaySpreadOpenOdds;
    public XSSFWorkbook buildExcel(XSSFWorkbook sportDataWorkbook, String dataEventID, String dataGame , int eventIndex, Elements soupOddsElements, Elements nflElements)
    {
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        Date date = new Date();
        String time = (dateFormat.format(date));
        sportDataSheet = sportDataWorkbook.getSheet("Data");
        sportDataSheet.setDefaultColumnWidth(10);
        CellStyle leftStyle = sportDataWorkbook.createCellStyle();
        CellStyle centerStyle = sportDataWorkbook.createCellStyle();
        CellStyle myStyle = sportDataWorkbook.createCellStyle();
        XSSFCellStyle redStyle = sportDataWorkbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        sportDataSheet.setDefaultColumnStyle(0, leftStyle);
        sportDataSheet.setDefaultColumnStyle(1, centerStyle);
        sportDataSheet.createRow(eventIndex);
        thisMatchupDate = gameDatesMap.get(dataEventID);
        atsHome = atsHomesMap.get(dataEventID);
        atsAway = atsAwaysMap.get(dataEventID);
        ouOver = ouOversMap.get(dataEventID);
        ouUnder = ouUndersMap.get(dataEventID);
        XSSFCreationHelper createHelper = sportDataWorkbook.getCreationHelper();
        XSSFCellStyle cellStyle = sportDataWorkbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMMM dd, yyyy"));
        dataEventIdElements = nflElements.select("[data-event-id='" + dataEventID + "']");
        bet365DataGameElements = soupOddsElements.select("[data-book='bet365'][data-game='" + dataGame + "']");
        sportDataSheet.autoSizeColumn(0);//Time stamp e.g. 2022/08/14 20:28:42
        sportDataSheet.getRow(eventIndex).createCell(0);//Time stamp
        sportDataSheet.getRow(eventIndex).getCell(0).setCellStyle(leftStyle);
        sportDataSheet.getRow(0).getCell(0).setCellValue(time);
        thisMatchupDate = dataEventIdElements.select(".cmg_matchup_header_date").text().split(",")[1];
        sportDataSheet.autoSizeColumn(1);//Matchup up date e.g. 2022-09-11
        sportDataSheet.getRow(eventIndex).createCell(1);
        sportDataSheet.getRow(eventIndex).getCell(1).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(1).setCellValue(thisMatchupDate);
        sportDataSheet.autoSizeColumn(2);
        sportDataSheet.getRow(eventIndex).createCell(2);//Season e.g. 2022
        sportDataSheet.getRow(eventIndex).getCell(2).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(2).setCellValue(season);
        sportDataSheet.autoSizeColumn(3);
        sportDataSheet.getRow(eventIndex).createCell(3);//Week number e.g.4
        sportDataSheet.getRow(eventIndex).getCell(3).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(3).setCellValue(weekNumber);
        sportDataSheet.autoSizeColumn(4);
        String calendarMonth = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(" ")[1];
        sportDataSheet.getRow(eventIndex).createCell(4);//Month e.g. Sep
        sportDataSheet.getRow(eventIndex).getCell(4).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(4).setCellValue(calendarMonth);
        sportDataSheet.autoSizeColumn(5);
        String calendarDay = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(",")[0];
        Cell dateCell = sportDataSheet.getRow(eventIndex).createCell(5);//Day of the week e.g. Monday
        sportDataSheet.getRow(eventIndex).getCell(5).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(5).setCellValue(calendarDay);
        sportDataSheet.autoSizeColumn(10);
        homeNickname = dataEventIdElements.attr("data-home-team-nickname-search");
        homeCity = dataEventIdElements.attr("data-home-team-city-search");
        sportDataSheet.getRow(eventIndex).createCell(10);//Home team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(eventIndex).getCell(10).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(10).setCellValue(homeCity + " " + homeNickname);

        sportDataSheet.autoSizeColumn(11);
        homeTeamShortName = dataEventIdElements.attr("data-home-team-shortname-search");//Home team abbreviation e.g. LAR
        sportDataSheet.getRow(eventIndex).createCell(11);
        sportDataSheet.getRow(eventIndex).getCell(11).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(11).setCellValue(homeTeamShortName);





        sportDataSheet.autoSizeColumn(13);//Data EventId/DataGame column M 13
        sportDataSheet.getRow(eventIndex).createCell(12);
        sportDataSheet.getRow(eventIndex).getCell(12).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(12).setCellValue(dataEventID + "/" + dataGame);




        sportDataSheet.autoSizeColumn(13);//Home spread open odds column N 14*******************
        homeSpreadOpenOdds = bet365DataGameElements.select("[data-type='spread']").text().split(" ")[6];
        sportDataSheet.getRow(eventIndex).createCell(13);
        sportDataSheet.getRow(eventIndex).getCell(13).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(13).setCellValue(homeSpreadOpenOdds);
        System.out.println("EB148========================HomeSpreadOpen => " + homeSpreadOpenOdds);

        sportDataSheet.autoSizeColumn(14);//bet365 Home spread close odds column O 15**************
        homeSpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__homeOdds .__decimal").text().split(" ")[0];
        sportDataSheet.getRow(eventIndex).createCell(14);
        sportDataSheet.getRow(eventIndex).getCell(14).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(14).setCellValue(homeSpreadCloseOdds);

        sportDataSheet.autoSizeColumn(18);//bet365 Home moneyline close odds column S 19**********
        homeMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__homeOdds .__american").text();
        sportDataSheet.getRow(eventIndex).createCell(18);
        sportDataSheet.getRow(eventIndex).getCell(18).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(18).setCellValue(homeMoneylineCloseOdds);
        sportDataSheet.autoSizeColumn(17);
        sportDataSheet.getRow(eventIndex).createCell(17);//MoneyLine Bet365 home odds, column R
        sportDataSheet.getRow(eventIndex).getCell(17).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(17).setCellValue(homeMoneyLineOddsMap.get(dataEventID));
        sportDataSheet.autoSizeColumn(25);
        awayNickname = dataEventIdElements.attr("data-away-team-nickname-search");
        awayCity = dataEventIdElements.attr("data-away-team-city-search");
        sportDataSheet.getRow(eventIndex).createCell(25);//Away team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(eventIndex).getCell(25).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(25).setCellValue(awayCity + " " + awayNickname);
        String homeTeamPlusNickname = homeCity + " " + homeNickname;
        String awayTeamPlusNickname = awayCity + " " + awayNickname;
        gameIdentifier = season + " - " + awayTeamPlusNickname + " @ " + homeTeamPlusNickname;
        sportDataSheet.getRow(eventIndex).getCell(0).setCellValue(gameIdentifier);//e.g. 2022 - Washington Football Team @ Dallas Cowboys
        sportDataSheet.autoSizeColumn(26);//Away team shortName e.g. LAR
        awayTeamShortName = dataEventIdElements.attr("data-away-team-shortname-search");
        sportDataSheet.getRow(eventIndex).createCell(26);
        sportDataSheet.getRow(eventIndex).getCell(26).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(26).setCellValue(awayTeamShortName);



        sportDataSheet.autoSizeColumn(28);//Away spread open Column AC 29*********
        awaySpreadOpenOdds =  bet365DataGameElements.select("[data-type='spread']").text().split(" ")[0];
        System.out.println("EB185 ............ ASO=> " + awaySpreadOpenOdds);
        sportDataSheet.getRow(eventIndex).createCell(28);
        sportDataSheet.getRow(eventIndex).getCell(28).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(28).setCellValue(awaySpreadOpenOdds);





        sportDataSheet.autoSizeColumn(29);//bet365 Away spread close odds column AD 30
        awaySpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__awayOdds div.__american").text().split(" ")[0];//********
        sportDataSheet.getRow(eventIndex).createCell(29);
        sportDataSheet.getRow(eventIndex).getCell(29).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(29).setCellValue(awaySpreadCloseOdds);

        sportDataSheet.autoSizeColumn(33);//bet365 Away moneyline close odds column AH 34*************
        awayMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__awayOdds .__american").text();
        sportDataSheet.getRow(eventIndex).createCell(33);
        sportDataSheet.getRow(eventIndex).getCell(33).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(33).setCellValue(awayMoneylineCloseOdds);
        sportDataSheet.autoSizeColumn(59);
        sportDataSheet.getRow(eventIndex).createCell(59);
        sportDataSheet.getRow(eventIndex).getCell(59).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(59).setCellValue(atsHome);
        sportDataSheet.autoSizeColumn(61);
        sportDataSheet.getRow(eventIndex).createCell(61);
        sportDataSheet.getRow(eventIndex).getCell(61).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(61).setCellValue(atsAway);
        sportDataSheet.autoSizeColumn(64);
        sportDataSheet.getRow(eventIndex).createCell(64);
        sportDataSheet.getRow(eventIndex).getCell(64).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(64).setCellValue(ouOver);
        sportDataSheet.autoSizeColumn(66);
        sportDataSheet.getRow(eventIndex).createCell(66);
        sportDataSheet.getRow(eventIndex).getCell(66).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(66).setCellValue(ouUnder);
        return sportDataWorkbook;
    }
    public void setGameDatesMap(HashMap<String, String> gameDatesMap)
    {
        this.gameDatesMap = gameDatesMap;
    }
    public void setAtsHomesMap(HashMap<String, String> atsHomes)
    {
        this.atsHomesMap = atsHomes;
    }
    public void setAtsAwaysMap(HashMap<String, String> atsAwayMap)
    {
        this.atsAwaysMap = atsAwayMap;
    }
    public void setOuOversMap(HashMap<String, String> ouOversMap)
    {
        this.ouOversMap = ouOversMap;
    }
    public void setOuUndersMap(HashMap<String, String> ouUndersMap)
    {
        this.ouUndersMap = ouUndersMap;
    }
    public void setGameIdentifier(String gameIdentifier)
    {
        this.gameIdentifier = gameIdentifier;
    }
    public void setHomeMLOddsMap(HashMap<String, String> homeMLOddsMap)
    {
        this.homeMLOddsMap = homeMLOddsMap;
    }
    public void setMoneyLineOdds(String moneyLineOdds, String dataEventId)
    {
    }
    public void setSpreadOdds(String spreadOdds, String dataEventId)
    {
        String[] spreadOddsArray = spreadOdds.split(" ");
        if (spreadOddsArray.length > 0)
        {
            awaySpreadCloseOdds = spreadOddsArray[0];
            awaySpreadOddsMap.put(dataEventId, awayMoneyLineOdds);
            homeSpreadOdds = spreadOddsArray[1];
            homeSpreadOddsMap.put(dataEventId, homeMoneyLineOdds);

        }
    }
    public void setSeason(String season)
    {
        this.season = season;
    }
    public void setWeekNumber(String weekNumber)
    {
        this.weekNumber = weekNumber;
    }
    public String getGameIdentifier()
    {
        return gameIdentifier;
    }
    public String getThisMatchupDate()
    {
        return thisMatchupDate;
    }
    public Elements getBet365DataGameElements()
    {
        return bet365DataGameElements;
    }
}

