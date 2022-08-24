package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version 220823
 * Builds data event id array and calendar date array
 *******************************************************************/
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.HashMap;
public class DataCollector
{
    private static HashMap<String, String> bet365HomeTeamOdds = new HashMap<>();
    private static HashMap<String, String> bet365AwayTeamOdds = new HashMap<>();
    private static HashMap<String, String> bet365Odds = new HashMap<>();
    private static ArrayList<String> thisWeekMatchuplist = new ArrayList<>();
    private static ArrayList<String> homeAmericanOddsArray = new ArrayList<>();
    private static HashMap<String, String> homeAmericanOddsMap = new HashMap<>();
    private static ArrayList<String> homeDecimalOddsArray = new ArrayList<>();
    private static HashMap<String, String> homeDecimalOddsMap = new HashMap<>();
    private static ArrayList<String> homeFractionalOddsArray = new ArrayList<>();
    private static HashMap<String, String> homeFractionalOddsMap = new HashMap<>();
    private static ArrayList<String> awayAmericanOddsArray = new ArrayList<>();
    private static HashMap<String, String> awayAmericanOddsMap = new HashMap<>();
    private static ArrayList<String> awayDecimalOddsArray = new ArrayList<>();
    private static HashMap<String, String> awayMLoddsMap = new HashMap<>();
    private static HashMap<String, String> homeMLoddsMap = new HashMap<>();
    private static ArrayList<String> awayFractionalOddsArray = new ArrayList<>();
    private static HashMap<String, String> awayFractionalOddsMap = new HashMap<>();
    private HashMap<String, String> mlHomeOdds = new HashMap<String, String>();
    private HashMap<String, String> mlAwayOdds = new HashMap<String, String>();
    private String dataEventId;
    private String MLhomeOdds;
    private String MLawayOdds;
    private String homeTeamNickname;//e.g. Browns...data-home-team-nickname-search
    private String awayTeamNickname;//e.g Texans...data-away-team-nickname-search
    private String awayTeamFullName;//e.g. Cleveland...data-home-team-fullname-search
    private String homeTeamFullName;//e.g Houston...data-home-team-fullname-search
    private String awayTeamCompleteName;//e.g. Kansas City Chiefs
    private static String homeTeamCompleteName;//e.g Houston Texans
    private String gameIdentifier;//e.g 2020 - Houston Texans @ Kansas City Chiefs
    private String awayTeamScore;
    private String homeTeamScore;
    private String gameDate;
    private String awayTeamCity;
    private String homeTeamCity;
    private String thisWeek;
    private String thisSeason = "2021";
    private ArrayList<String> thisWeekGameDates = new ArrayList<String>();
    private ArrayList<String> thisGameWeekNumbers = new ArrayList<String>();
    private ArrayList<String> thisWeekHomeTeamScores = new ArrayList<String>();
    private ArrayList<String> thisWeekAwayTeamScores = new ArrayList<String>();
    private ArrayList<String> thisWeekHomeTeams = new ArrayList<String>();
    private ArrayList<String> atsHomes = new ArrayList<String>();
    private ArrayList<String> thisWeekAwayTeams = new ArrayList<String>();
    private HashMap<String, String> gameDatesMap = new HashMap<>();
    private HashMap<String, String> gameIdentifierMap = new HashMap<>();
    private HashMap<String, String> homeFullNameMap = new HashMap<>();
    private HashMap<String, String> awayFullNameMap = new HashMap<>();
    private HashMap<String, String> homeShortNameMap = new HashMap<>();
    private HashMap<String, String> awayShortNameMap = new HashMap<>();
    private HashMap<String, String> atsHomesMap = new HashMap<>();
    private HashMap<String, String> atsAwaysMap = new HashMap<>();
    private HashMap<String, String> ouUndersMap = new HashMap<>();
    private HashMap<String, String> ouOversMap = new HashMap<>();
    private HashMap<String, String> cityNameMap = new HashMap<>();
    private HashMap<String, String> idXref = new HashMap<>();
    private String[] bet365OddsArray = new String[6];
    public HashMap<String, String> getAwayMLoddsMap() {return awayMLoddsMap;}
    public HashMap<String, String> getHomeMLoddsMap() {return awayMLoddsMap;}
    public void collectTeamInfo(Elements thisWeekElements)//From covers.com website for this week's matchups
    {
        for (Element e : thisWeekElements)//Build week matchup IDs array
        {
            homeTeamFullName = e.attr("data-home-team-fullname-search");//e.g. Houston...correcting for different city/name usage
            homeTeamNickname = e.attr("data-home-team-nickname-search");//e.g. Texans
            homeTeamCity = e.attr("data-home-team-city-search");
            homeTeamCity = cityNameMap.get(homeTeamCity);
            homeTeamCompleteName = homeTeamCity + " " + homeTeamNickname;
            awayTeamFullName = e.attr("data-away-team-fullname-search");//e.g. Dallas
            awayTeamNickname = e.attr("data-away-team-nickname-search");//e.g. Cowboys
            awayTeamCity = e.attr("data-away-team-city-search");
            awayTeamCity = cityNameMap.get(awayTeamCity);
            awayTeamCompleteName = awayTeamCity + " " + awayTeamNickname;
            gameIdentifier = thisSeason + " - " + awayTeamCompleteName + " @ " + homeTeamCompleteName;
            dataEventId = e.attr("data-event-id");
            String[] gameDateTime = e.attr("data-game-date").split(" ");
            gameDate = gameDateTime[0];
            awayTeamScore = e.attr("data-away-score");
            thisWeek = e.attr("data-competition-type");
            thisWeekGameDates.add(gameDate);
            gameDatesMap.put(dataEventId, gameDate);
            gameIdentifierMap.put(dataEventId, gameIdentifier);
            thisWeekHomeTeams.add(homeTeamCompleteName);
            thisWeekAwayTeams.add(awayTeamCompleteName);
            homeFullNameMap.put(dataEventId, homeTeamFullName);
            awayFullNameMap.put(dataEventId, awayTeamFullName);
            thisWeekHomeTeamScores.add(homeTeamScore);
            thisWeekAwayTeamScores.add((awayTeamScore));
            thisGameWeekNumbers.add(thisWeek);
            String awayShortName = e.attr("data-away-team-shortname-search");//Away team
            awayShortNameMap.put(dataEventId, awayShortName);
            String homeShortName = e.attr("data-home-team-shortname-search");//Home team
            homeShortNameMap.put(dataEventId, homeShortName);
        }
    }
    public void collectConsensusData(Elements thisMatchupConsensus, String thisMatchupID)
    {
        this.dataEventId = thisMatchupID;
        String ouOver = null;
        String ouUnder = null;
        String atsHome = null;
        String atsAway = null;
        Elements rightConsensus = thisMatchupConsensus.select(".covers-CoversConsensusDetailsTable-finalWagersright");//Home/Under
        Elements leftConsensus = thisMatchupConsensus.select(".covers-CoversConsensusDetailsTable-finalWagersleft");//Away/Over
        try//To catch missing consensus data due to delayed or cancelled game
        {
            ouUnder = rightConsensus.select("div").get(1).text();
            ouOver = leftConsensus.select("div").get(1).text();
            atsHome = leftConsensus.select("div").get(0).text();
            atsAway = rightConsensus.select("div").get(0).text();
        }
        catch (Exception e)
        {
            System.out.println("DC121 DataCollector, no consensus data for " + gameIdentifier);
        }
        ouOversMap.put(thisMatchupID, ouOver);
        ouUndersMap.put(thisMatchupID, ouUnder);
        atsHomesMap.put(thisMatchupID, atsAway);
        atsAwaysMap.put(thisMatchupID, atsHome);
    }
    public String collectMoneylineOdds(Elements oddsElements, HashMap<String, String> xRefMap, String dataEventId)
    {
//        String dataGame = xRefMap.get(dataEventId);//e.g. data-event-id...83480 => data-game...244451
//        String[] moneyLineOddsArray = oddsElements.select("*").text().split(" ");
//        String moneyLineHomeOdds = moneyLineOddsArray[0];
//        String moneyLineAwayOdds = moneyLineOddsArray[3];
//        String moneylineOddsString = moneyLineAwayOdds + " " + moneyLineHomeOdds;
//        System.out.println(moneylineOddsString);
//        return moneylineOddsString;//awayOdds space homeOdds
        return "0";
    }
    public HashMap<String, String> getHomeFullNameMap()
    {
        return homeFullNameMap;
    }
    public HashMap<String, String> getAwayFullNameMap()
    {
        return awayFullNameMap;
    }
    public HashMap<String, String> getGameDatesMap()
    {
        return gameDatesMap;
    }
    public HashMap<String, String> getAtsHomesMap()
    {
        return atsHomesMap;
    }
    public HashMap<String, String> getAtsAwaysMap()
    {
        return atsAwaysMap;
    }
    public HashMap<String, String> getOuOversMap()
    {
        return ouOversMap;
    }
    public HashMap<String, String> getOuUndersMap()
    {
        return ouUndersMap;
    }
    public HashMap<String, String> getGameIdentifierMap(){return gameIdentifierMap;}
    public void setThisSeason(String thisSeason)
    {
        this.thisSeason = thisSeason;
    }
    public String getAwayTeamCompleteName()
    {
        return awayTeamCompleteName;
    }
    public String getHomeTeamCompleteName()
    {
        return homeTeamCompleteName;
    }
    public void setCityNameMap(HashMap<String, String> cityNameMap)
    {
        this.cityNameMap = cityNameMap;
    }
}
