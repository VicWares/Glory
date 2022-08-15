package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220814B
 *******************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Spliterator;
public class ExcelBuilder
{
    private String season;
    private String weekNumber;
    private String ouUnder;
    private String ouOver;
    private String homeTeam;
    private String awayTeam;
    private String thisMatchupDate;
    private String atsHome;
    private String atsAway;
    private String completeHomeTeamName;
    private String completeAwayTeamName;
    private String gameIdentifier;
    private String awayMoneyLineOdds;
    private String homeMoneyLineOdds;
    private String awaySpreadOdds;
    private String homeSpreadOdds;
    private String awayNickname;
    private String homeNickname;
    private HashMap<String, String> homeTeamsMap = new HashMap<>();
    private HashMap<String, String> awayTeamsMap = new HashMap<>();
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
    public XSSFWorkbook buildExcel(XSSFWorkbook sportDataWorkbook, String dataEventID, int eventIndex, String gameIdentifier, Elements nflElements)
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
        sportDataSheet.setColumnWidth(1, 25 * 256);
        homeTeam = homeTeamsMap.get(dataEventID);
        awayTeam = awayTeamsMap.get(dataEventID);
        thisMatchupDate = gameDatesMap.get(dataEventID);
        atsHome = atsHomesMap.get(dataEventID);
        atsAway = atsAwaysMap.get(dataEventID);
        ouOver = ouOversMap.get(dataEventID);
        ouUnder = ouUndersMap.get(dataEventID);

        XSSFCreationHelper createHelper = sportDataWorkbook.getCreationHelper();
        XSSFCellStyle cellStyle         = sportDataWorkbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMMM dd, yyyy"));

        Elements dataEventIdElements = nflElements.select("[data-event-id='" + dataEventID + "']");//All elements for this matchup
        sportDataSheet.autoSizeColumn(0);
        sportDataSheet.getRow(eventIndex).createCell(0);//Time stamp
        sportDataSheet.getRow(eventIndex).getCell(0).setCellStyle(leftStyle);
        sportDataSheet.getRow(0).getCell(0).setCellValue(time);

        String homeTeamPlusNickname = homeTeam + " " + homeNickname;
        String awayTeamPlusNickname = awayTeam + " " + awayNickname;
        gameIdentifier = season + " - " + awayTeamPlusNickname + " @ " + homeTeamPlusNickname;
        sportDataSheet.getRow(eventIndex).getCell(0).setCellValue(gameIdentifier);//e.g. 2022 - Washington Football Team @ Dallas Cowboys

        sportDataSheet.autoSizeColumn(1);
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
        dataEventIdElements.select(".cmg_matchup_header_date");
        String calendarMonth = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(" ")[1];
        sportDataSheet.getRow(eventIndex).createCell(4);//Month e.g. Sep
        sportDataSheet.getRow(eventIndex).getCell(4).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(4).setCellValue(calendarMonth);

        sportDataSheet.autoSizeColumn(5);
        dataEventIdElements.select(".cmg_matchup_header_date");
        String calendarDay = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(",")[0];
        Cell dateCell =  sportDataSheet.getRow(eventIndex).createCell(5);//Day of the week e.g. Monday
        sportDataSheet.getRow(eventIndex).getCell(5).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(5).setCellValue(calendarDay);

        sportDataSheet.autoSizeColumn(10);
        homeNickname =  dataEventIdElements.attr("data-home-team-nickname-search");
        sportDataSheet.getRow(eventIndex).createCell(10);//Home team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(eventIndex).getCell(10).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(10).setCellValue(homeTeam + " " + homeNickname);

        sportDataSheet.autoSizeColumn(12);
        sportDataSheet.getRow(eventIndex).createCell(12);//Spread home odds, column M
        sportDataSheet.getRow(eventIndex).getCell(12).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(12).setCellValue(homeSpreadOddsMap.get(dataEventID));

        sportDataSheet.autoSizeColumn(17);
        sportDataSheet.getRow(eventIndex).createCell(17);//MoneyLine Bet365 home odds, column R
        sportDataSheet.getRow(eventIndex).getCell(17).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(17).setCellValue(homeMoneyLineOddsMap.get(dataEventID));

        sportDataSheet.autoSizeColumn(25);
        awayNickname =  dataEventIdElements.attr("data-away-team-nickname-search");
        sportDataSheet.getRow(eventIndex).createCell(25);//Away team + nickname
        sportDataSheet.getRow(eventIndex).getCell(25).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(25).setCellValue(awayTeam + " " + awayNickname);

        sportDataSheet.autoSizeColumn(26);
        sportDataSheet.getRow(eventIndex).createCell(26);//Spread away odds, column AA
        sportDataSheet.getRow(eventIndex).getCell(26).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(26).setCellValue(awaySpreadOddsMap.get(dataEventID));

        sportDataSheet.autoSizeColumn(31);
        sportDataSheet.getRow(eventIndex).createCell(31);//MoneyLine Bet365 away odds, column AF
        sportDataSheet.getRow(eventIndex).getCell(31).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(31).setCellValue(awayMoneyLineOddsMap.get(dataEventID));

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
    public void setHomeTeamsMap(HashMap<String, String> homeTeamsMap)
    {
        this.homeTeamsMap = homeTeamsMap;
    }
    public void setThisWeekAwayTeamsMap(HashMap<String, String> thisWeekAwayTeamsMap)
    {
        this.awayTeamsMap = thisWeekAwayTeamsMap;
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
    public void setCompleteHomeTeamName(String completeHomeTeamName)
    {
        this.completeHomeTeamName = completeHomeTeamName;
    }
    public void setCompleteAwayTeamName(String completeAwayTeamName)
    {
        this.completeAwayTeamName = completeAwayTeamName;
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
            awaySpreadOdds = spreadOddsArray[0];
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
}

