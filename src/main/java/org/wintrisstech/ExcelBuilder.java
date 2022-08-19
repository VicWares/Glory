package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220819A
 *******************************************************************/
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
    private String matchupDate;
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
    private HashMap<String, String> cityNameMap = new HashMap<>();
    private String awayCity;
    private String homeTeamShortName;
    private String awayTeamShortName;
    private String awayTeamSpreadOpenOdds;
    private String homeMoneylineCloseOdds;
    private String homeSpreadOpenOdds;
    private String homeSpreadCloseOdds;
    private String awayMoneylineCloseOdds;
    private Elements dataEventIdElements;
    private Elements bet365DataGameElements;
    private String awaySpreadOpenOdds;
    private String awaySpreadConsensus;
    private String homeSpreadConsensus;
    private Elements consensusElements;
    private String calendarMonth;
    private String calendarDay;
    private String consensusSpreadAway;
    private String consensusSpreadHome;
    private String consnsusOUunderAway;
    private String consnsusOUunderHome;
    private String consnsusOverUunderHome;
    private String consnsusOverUunderAway;
    private String consensusMoneyLeaderSpreadHome;
    private String consensusMoneyLeaderSpreadAway;
    private String consnsusMoneyLeaderOverUnderAway;
    private String consnsusMoneyLeaderOverUnderHome;
    public XSSFWorkbook buildExcel(XSSFWorkbook sportDataWorkbook, String dataEventID, String dataGame , int excelRowIndex, Elements soupOddsElements, Elements nflElements)
    {
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        Date date = new Date();
        String time = (dateFormat.format(date));
        sportDataSheet = sportDataWorkbook.getSheet("Data");
        sportDataSheet.setDefaultColumnWidth(10);
        XSSFCellStyle redStyle = sportDataWorkbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        sportDataSheet.createRow(excelRowIndex);
        matchupDate = gameDatesMap.get(dataEventID);

        XSSFCreationHelper createHelper = sportDataWorkbook.getCreationHelper();
        XSSFCellStyle cellStyle = sportDataWorkbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMMM dd, yyyy"));
        dataEventIdElements = nflElements.select("[data-event-id='" + dataEventID + "']");
        bet365DataGameElements = soupOddsElements.select("[data-book='bet365'][data-game='" + dataGame + "']");

        setSportDataSheet(0, 0, time);//Time stamp row/cel 00
        gameIdentifier = season + " - " + awayCity + " " + awayNickname +" @ " + homeCity + " " + homeNickname;//Game identifier e.g. 2022 - Buffalo Bills @ Los Angeles Rams Column A 1
        setSportDataSheet(excelRowIndex, 0, gameIdentifier);//Game Identifier e.g. 2022 - Pittsburgh Steelers @ Jacksonville Jaguars column A 1
        setSportDataSheet(excelRowIndex, 1,matchupDate);//Matchup up date e.g. Aug. 20 Column B2
        setSportDataSheet(excelRowIndex, 2,season);///Season e.g. 2022 Column C 3
        setSportDataSheet(excelRowIndex, 3,weekNumber);//Week number e.g. 4 Column D 4
        calendarMonth = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(" ")[1];
        setSportDataSheet(excelRowIndex, 4,calendarMonth);//Calendar month e.g. Sep. Column E 5
        calendarDay = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(",")[0];
        setSportDataSheet(excelRowIndex, 5, calendarDay);// Calendar day e.g. Sunday Column F 6
        homeCity = dataEventIdElements.attr("data-home-team-city-search");
        homeCity = cityNameMap.get(homeCity);//Correct for bad city name
        homeNickname = dataEventIdElements.attr("data-home-team-nickname-search");
        String homeCityPlusNickname = homeCity + " " + homeNickname;
        setSportDataSheet(excelRowIndex, 10, homeCityPlusNickname);// Home team city + nickname e.g. Seattle Seahawks Column K 11
        homeTeamShortName = dataEventIdElements.attr("data-home-team-shortname-search");//Home team abbreviation e.g. LAR
        setSportDataSheet(excelRowIndex, 11, homeTeamShortName);// Home team short name e.g. DAL Column L 12
        homeSpreadOpenOdds = bet365DataGameElements.select("[data-type='spread']").text().split(" ")[6];
        setSportDataSheet(excelRowIndex, 13, homeSpreadOpenOdds);//Home spread open odds e.g. +3.5 column N 14
        homeSpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__homeOdds .__decimal").text().split(" ")[0];
        homeMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__homeOdds .__american").text();
        setSportDataSheet(excelRowIndex, 18, homeMoneylineCloseOdds);//bet365 Home moneyline close odds e.g +235 column S 19
        awayCity = dataEventIdElements.attr("data-away-team-city-search");
        awayCity = cityNameMap.get(awayCity);//Correct for weird Covers city names
        awayNickname = dataEventIdElements.attr("data-away-team-nickname-search");
        String awayCityPlusNickname = awayCity + " " + awayNickname;
        setSportDataSheet(excelRowIndex, 25, awayCityPlusNickname);//Away team city + nickname e.g. Jacksonville Jaguars Column Z 26
        awayTeamShortName = dataEventIdElements.attr("data-away-team-shortname-search");
        setSportDataSheet(excelRowIndex, 26, awayTeamShortName);//Away team shortName e.g. LAR Column AA 27
        awaySpreadOpenOdds =  bet365DataGameElements.select("[data-type='spread']").text().split(" ")[0];
        setSportDataSheet(excelRowIndex, 28, awaySpreadOpenOdds);//Away spread open Column AC 29
        awaySpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__awayOdds div.__american").text().split(" ")[0];//********
        setSportDataSheet(excelRowIndex, 29, awaySpreadCloseOdds);//bet365 Away spread close odds column AD 30
        awayMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__awayOdds .__american").text();
        setSportDataSheet(excelRowIndex, 30, awayMoneylineCloseOdds);//bet365 Away spread close odds column AD 30
        consensusSpreadAway = consensusElements.select("div.covers-CoversConsensusDetailsTable-row:nth-child(11) > div:nth-child(3) > div:nth-child(1)").text();
        setSportDataSheet(excelRowIndex, 64, consensusSpreadAway);//Consensus spread away column BM 65
        setSportDataSheet(excelRowIndex, 66, consensusSpreadHome);//Consesus Spread home column BO 67
        setSportDataSheet(excelRowIndex, 67, consensusMoneyLeaderSpreadAway);//Consesus Money Leaders Spread home column BP 68
        setSportDataSheet(excelRowIndex, 73, consnsusMoneyLeaderOverUnderAway);//Consensus Money Leader o/u under home column column BV 74
        setSportDataSheet(excelRowIndex, 74, consnsusMoneyLeaderOverUnderHome);//Consensus o/u under home column column BW 75
        homeSpreadConsensus = consensusElements.select("div.covers-CoversConsensusDetailsTable-homeWagers").text();// Homespread consensus BQ 69
        System.out.println("EB130..............home city => " + homeCity + " ...........away city => " + awayCity + " ............matchup date => " + matchupDate + " " + season);
        return sportDataWorkbook;
    }
    public void setSportDataSheet(int rowIndex, int column, String columnEntry)
    {
        sportDataSheet.autoSizeColumn(column);
        sportDataSheet.autoSizeColumn(column);//Zero based column numbers
        sportDataSheet.getRow(rowIndex).createCell(column);
        sportDataSheet.getRow(rowIndex).getCell(column).setCellValue(columnEntry);
    }
    public void setGameIdentifier(String gameIdentifier)
    {
        this.gameIdentifier = gameIdentifier;
    }
    public void setSeason(String season)
    {
        this.season = season;
    }
    public void setWeekNumber(String weekNumber) {this.weekNumber = weekNumber;}
    public void setCityNameMap(HashMap<String, String> cityNameMap) {this.cityNameMap = cityNameMap;}
    public void setConsensusElements(Elements consensusElements) {this.consensusElements = consensusElements;}
}

