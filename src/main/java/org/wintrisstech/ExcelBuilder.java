package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220818A
 *******************************************************************/
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
    public XSSFWorkbook buildExcel(XSSFWorkbook sportDataWorkbook, String dataEventID, String dataGame , int excelRowIndex, Elements soupOddsElements, Elements nflElements)
    {
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        Date date = new Date();
        String time = (dateFormat.format(date));
        sportDataSheet = sportDataWorkbook.getSheet("Data");
        sportDataSheet.setDefaultColumnWidth(10);
        CellStyle centerStyle = sportDataWorkbook.createCellStyle();
        CellStyle myStyle = sportDataWorkbook.createCellStyle();
        XSSFCellStyle redStyle = sportDataWorkbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        sportDataSheet.setDefaultColumnStyle(1, centerStyle);
        sportDataSheet.createRow(excelRowIndex);
        matchupDate = gameDatesMap.get(dataEventID);
//        atsHome = atsHomesMap.get(dataEventID);
//        atsAway = atsAwaysMap.get(dataEventID);
//        ouOver = ouOversMap.get(dataEventID);
//        ouUnder = ouUndersMap.get(dataEventID);
        XSSFCreationHelper createHelper = sportDataWorkbook.getCreationHelper();
        XSSFCellStyle cellStyle = sportDataWorkbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMMM dd, yyyy"));
        dataEventIdElements = nflElements.select("[data-event-id='" + dataEventID + "']");
        bet365DataGameElements = soupOddsElements.select("[data-book='bet365'][data-game='" + dataGame + "']");

        sportDataSheet.autoSizeColumn(0);//Time stamp e.g. 2022/08/14 20:28:42 Column A 1 Row 0 only
        sportDataSheet.getRow(excelRowIndex).createCell(0);//Time stamp
        sportDataSheet.getRow(excelRowIndex).getCell(0).setCellStyle(centerStyle);
        sportDataSheet.getRow(0).getCell(0).setCellValue(time);

        gameIdentifier = season + " - " + awayCity + " " + awayNickname +" @ " + homeCity + " " + homeNickname;//Game identifier e.g. 2022 - Buffalo Bills @ Los Angeles Rams Column A 1
        sportDataSheet.autoSizeColumn(0);//Matchup up date e.g. Aug. 20 Column B2
        sportDataSheet.getRow(excelRowIndex).createCell(0);
        sportDataSheet.getRow(excelRowIndex).getCell(0).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(0).setCellValue(gameIdentifier);

        matchupDate = dataEventIdElements.select(".cmg_matchup_header_date").text().split(",")[1];
        sportDataSheet.autoSizeColumn(1);//Matchup up date e.g. Aug. 20 Column B2
        sportDataSheet.getRow(excelRowIndex).createCell(1);
        sportDataSheet.getRow(excelRowIndex).getCell(1).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(1).setCellValue(matchupDate);

        sportDataSheet.autoSizeColumn(2);//Season e.g. 2022 Column C 3
        sportDataSheet.getRow(excelRowIndex).createCell(2);
        sportDataSheet.getRow(excelRowIndex).getCell(2).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(2).setCellValue(season);

        sportDataSheet.autoSizeColumn(3);//Week number e.g.4 Column D 4
        sportDataSheet.getRow(excelRowIndex).createCell(3);
        sportDataSheet.getRow(excelRowIndex).getCell(3).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(3).setCellValue(weekNumber);

        sportDataSheet.autoSizeColumn(4);//Calendar month e.g. Sep. Column E 5
        String calendarMonth = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(" ")[1];
        sportDataSheet.getRow(excelRowIndex).createCell(4);//Month e.g. Sep
        sportDataSheet.getRow(excelRowIndex).getCell(4).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(4).setCellValue(calendarMonth);

        sportDataSheet.autoSizeColumn(5);// Calendar day e.g. Sunday Column F 6
        String calendarDay = dataEventIdElements.select("div.cmg_matchup_header_date").text().split(",")[0];
        sportDataSheet.getRow(excelRowIndex).createCell(5);//Day of the week e.g. Monday
        sportDataSheet.getRow(excelRowIndex).getCell(5).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(5).setCellValue(calendarDay);

        sportDataSheet.autoSizeColumn(10);// Home team city + nickname e.g. Seattle Seahawks Column K 11
        homeCity = dataEventIdElements.attr("data-home-team-city-search");
        homeCity = cityNameMap.get(homeCity);
        System.out.println("EB128..............home city => " + homeCity + " ...........away city => " + awayCity + " ............matchup date => " + matchupDate + " " + season);
        homeNickname = dataEventIdElements.attr("data-home-team-nickname-search");
        sportDataSheet.getRow(excelRowIndex).createCell(10);//Home team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(excelRowIndex).getCell(10).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(10).setCellValue(homeCity + " " + homeNickname);

        sportDataSheet.autoSizeColumn(11);// Team short name e.g. DAL Column L 12
        homeTeamShortName = dataEventIdElements.attr("data-home-team-shortname-search");//Home team abbreviation e.g. LAR
        sportDataSheet.getRow(excelRowIndex).createCell(11);
        sportDataSheet.getRow(excelRowIndex).getCell(11).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(11).setCellValue(homeTeamShortName);

        sportDataSheet.autoSizeColumn(13);//Data EventId/DataGame column M 13
        sportDataSheet.getRow(excelRowIndex).createCell(12);
        sportDataSheet.getRow(excelRowIndex).getCell(12).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(12).setCellValue(dataEventID + "/" + dataGame);

        sportDataSheet.autoSizeColumn(13);//Home spread open odds e.g. +3.5 column N 14
       // homeSpreadOpenOdds = bet365DataGameElements.select("[data-type='spread']").text().split(" ")[6];
        sportDataSheet.getRow(excelRowIndex).createCell(13);
        sportDataSheet.getRow(excelRowIndex).getCell(13).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(13).setCellValue(homeSpreadOpenOdds);

        sportDataSheet.autoSizeColumn(14);//bet365 Home spread close odds e.g. -2.5 column O 15
        homeSpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__homeOdds .__decimal").text().split(" ")[0];
        sportDataSheet.getRow(excelRowIndex).createCell(14);
        sportDataSheet.getRow(excelRowIndex).getCell(14).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(14).setCellValue(homeSpreadCloseOdds);

        sportDataSheet.autoSizeColumn(18);//bet365 Home moneyline close odds e.g +235 column S 19
        homeMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__homeOdds .__american").text();
        sportDataSheet.getRow(excelRowIndex).createCell(18);
        sportDataSheet.getRow(excelRowIndex).getCell(18).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(18).setCellValue(homeMoneylineCloseOdds);

        sportDataSheet.autoSizeColumn(25);//Away team city + nickname e.g. Jacksonville Jaguars Column Z 26
        awayCity = dataEventIdElements.attr("data-away-team-city-search");
        awayCity = cityNameMap.get(awayCity);//Correct for weird Covers city names
        awayNickname = dataEventIdElements.attr("data-away-team-nickname-search");
        String awayCityPlusNickname = awayCity + " " + awayNickname;
        sportDataSheet.getRow(excelRowIndex).createCell(25);//Away team + nickname e.g. Dallas Coyboys
        sportDataSheet.getRow(excelRowIndex).getCell(25).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(25).setCellValue(awayCityPlusNickname);

        sportDataSheet.autoSizeColumn(26);//Away team shortName e.g. LAR Column AA 27
        awayTeamShortName = dataEventIdElements.attr("data-away-team-shortname-search");
        sportDataSheet.getRow(excelRowIndex).createCell(26);
        sportDataSheet.getRow(excelRowIndex).getCell(26).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(26).setCellValue(awayTeamShortName);

        sportDataSheet.autoSizeColumn(28);//Away spread open Column AC 29*********
        awaySpreadOpenOdds =  bet365DataGameElements.select("[data-type='spread']").text().split(" ")[0];
        sportDataSheet.getRow(excelRowIndex).createCell(28);
        sportDataSheet.getRow(excelRowIndex).getCell(28).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(28).setCellValue(awaySpreadOpenOdds);

        sportDataSheet.autoSizeColumn(29);//bet365 Away spread close odds column AD 30
        awaySpreadCloseOdds = bet365DataGameElements.select("[data-type='spread'] div.__awayOdds div.__american").text().split(" ")[0];//********
        sportDataSheet.getRow(excelRowIndex).createCell(29);
        sportDataSheet.getRow(excelRowIndex).getCell(29).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(29).setCellValue(awaySpreadCloseOdds);

        sportDataSheet.autoSizeColumn(33);//bet365 Away moneyline close odds column AH 34
        awayMoneylineCloseOdds = bet365DataGameElements.select("[data-type='moneyline'] .__awayOdds .__american").text();
        sportDataSheet.getRow(excelRowIndex).createCell(33);
        sportDataSheet.getRow(excelRowIndex).getCell(33).setCellStyle(centerStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(33).setCellValue(awayMoneylineCloseOdds);

        sportDataSheet.autoSizeColumn(59);
        sportDataSheet.getRow(excelRowIndex).createCell(59);
        sportDataSheet.getRow(excelRowIndex).getCell(59).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(59).setCellValue(atsHome);

        sportDataSheet.autoSizeColumn(61);
        sportDataSheet.getRow(excelRowIndex).createCell(61);
        sportDataSheet.getRow(excelRowIndex).getCell(61).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(61).setCellValue(atsAway);

        sportDataSheet.autoSizeColumn(64);
        sportDataSheet.getRow(excelRowIndex).createCell(64);
        sportDataSheet.getRow(excelRowIndex).getCell(64).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(64).setCellValue(ouOver);

        sportDataSheet.autoSizeColumn(66);
        sportDataSheet.getRow(excelRowIndex).createCell(66);
        sportDataSheet.getRow(excelRowIndex).getCell(66).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(66).setCellValue(ouUnder);

        sportDataSheet.autoSizeColumn(67);// Away spread consensus BP 68
        awaySpreadConsensus = consensusElements.select("div.covers-CoversConsensusDetailsTable-away div.covers-CoversConsensusDetailsTable-awayWagers").text();
        System.out.println("awayCponsensus.................." + awaySpreadConsensus);
        sportDataSheet.getRow(excelRowIndex).createCell(67);
        sportDataSheet.getRow(excelRowIndex).getCell(67).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(67).setCellValue(ouUnder);

        sportDataSheet.autoSizeColumn(68);// Homespread consensus BQ 69
        homeSpreadConsensus = consensusElements.select("div.covers-CoversConsensusDetailsTable-homeWagers").text();
        System.out.println("homeconsensus..................." + homeSpreadConsensus);
        sportDataSheet.getRow(excelRowIndex).createCell(68);
        sportDataSheet.getRow(excelRowIndex).getCell(68).setCellStyle(myStyle);
        sportDataSheet.getRow(excelRowIndex).getCell(68).setCellValue(ouUnder);

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
    public void setSeason(String season)
    {
        this.season = season;
    }
    public void setWeekNumber(String weekNumber) {this.weekNumber = weekNumber;}
    public void setCityNameMap(HashMap<String, String> cityNameMap) {this.cityNameMap = cityNameMap;}
    public void setConsensusElements(Elements consensusElements)
    {this.consensusElements = consensusElements;}
}

