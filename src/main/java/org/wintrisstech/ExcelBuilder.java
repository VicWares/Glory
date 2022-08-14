package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220814A
 *******************************************************************/
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Element;
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

        Elements dataEventIdElements = nflElements.select("[data-event-id='" + dataEventID + "']");

        sportDataSheet.getRow(eventIndex).createCell(0);
        sportDataSheet.getRow(eventIndex).getCell(0).setCellStyle(leftStyle);
        sportDataSheet.getRow(0).getCell(0).setCellValue(time);

        sportDataSheet.getRow(eventIndex).getCell(0).setCellValue(gameIdentifier);//e.g. 2021 - Washington Football Team @ Dallas Cowboys

        sportDataSheet.getRow(eventIndex).createCell(1);
        sportDataSheet.getRow(eventIndex).getCell(1).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(1).setCellValue(thisMatchupDate);

        sportDataSheet.getRow(eventIndex).createCell(2);//Season e.g. 2022
        sportDataSheet.getRow(eventIndex).getCell(2).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(2).setCellValue(season);

        sportDataSheet.getRow(eventIndex).createCell(3);//Week number e.g.4
        sportDataSheet.getRow(eventIndex).getCell(3).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(3).setCellValue(weekNumber);

        homeNickname =  dataEventIdElements.attr("data-home-team-nickname-search");
        sportDataSheet.getRow(eventIndex).createCell(10);//Home team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(eventIndex).getCell(10).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(10).setCellValue(homeTeam + " " + homeNickname);

        sportDataSheet.getRow(eventIndex).createCell(12);//Spread home odds, column M
        sportDataSheet.getRow(eventIndex).getCell(12).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(12).setCellValue(homeSpreadOddsMap.get(dataEventID));

        sportDataSheet.getRow(eventIndex).createCell(17);//MoneyLine Bet365 home odds, column R
        sportDataSheet.getRow(eventIndex).getCell(17).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(17).setCellValue(homeMoneyLineOddsMap.get(dataEventID));

        awayNickname =  dataEventIdElements.attr("data-away-team-nickname-search");
        sportDataSheet.getRow(eventIndex).createCell(25);//Away team + nickname
        sportDataSheet.getRow(eventIndex).getCell(25).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(25).setCellValue(awayTeam + " " + awayNickname);

        sportDataSheet.getRow(eventIndex).createCell(26);//Spread away odds, column AA
        sportDataSheet.getRow(eventIndex).getCell(26).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(26).setCellValue(awaySpreadOddsMap.get(dataEventID));

        sportDataSheet.getRow(eventIndex).createCell(31);//MoneyLine Bet365 away odds, column AF
        sportDataSheet.getRow(eventIndex).getCell(31).setCellStyle(centerStyle);
        sportDataSheet.getRow(eventIndex).getCell(31).setCellValue(awayMoneyLineOddsMap.get(dataEventID));

        sportDataSheet.getRow(eventIndex).createCell(59);
        sportDataSheet.getRow(eventIndex).getCell(59).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(59).setCellValue(atsHome);

        sportDataSheet.getRow(eventIndex).createCell(61);
        sportDataSheet.getRow(eventIndex).getCell(61).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(61).setCellValue(atsAway);

        sportDataSheet.getRow(eventIndex).createCell(64);
        sportDataSheet.getRow(eventIndex).getCell(64).setCellStyle(myStyle);
        sportDataSheet.getRow(eventIndex).getCell(64).setCellValue(ouOver);

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

