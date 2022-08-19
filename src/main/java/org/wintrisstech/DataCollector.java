package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220819A
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
    private ArrayList<String> thisWeekGameDates = new ArrayList<String>();
    private ArrayList<String> thisGameWeekNumbers = new ArrayList<String>();
    private ArrayList<String> thisWeekHomeTeamScores = new ArrayList<String>();
    private ArrayList<String> thisWeekAwayTeamScores = new ArrayList<String>();
    private ArrayList<String> thisWeekHomeTeams = new ArrayList<String>();
    private ArrayList<String> atsHomes = new ArrayList<String>();
    private ArrayList<String> thisWeekAwayTeams = new ArrayList<String>();
    private HashMap<String, String> gameDatesMap = new HashMap<>();
    private HashMap<String, String> gameIdentifierMap = new HashMap<>();
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
    ArrayList<String> oddStrings = new ArrayList<>();
    public HashMap<String, String> getAwayMLoddsMap()
    {
        return awayMLoddsMap;
    }
    public HashMap<String, String> getHomeMLoddsMap()
    {
        return awayMLoddsMap;
    }
    public void collectConsensusData(Elements Consensus, String MatchupID)
    {
        String ouOver = null;
        String ouUnder = null;
        String atsHome = null;
        String atsAway = null;
        Elements rightConsensus = Consensus.select(".covers-CoversConsensusDetailsTable-finalWagersright");//Home/Under
        Elements leftConsensus = Consensus.select(".covers-CoversConsensusDetailsTable-finalWagersleft");//Away/Over
        try//To catch missing consensus data due to delayed or cancelled game
        {
            ouUnder = rightConsensus.select("div").get(1).text();
            ouOver = leftConsensus.select("div").get(1).text();
            atsHome = leftConsensus.select("div").get(0).text();
            atsAway = rightConsensus.select("div").get(0).text();
        }
        catch (Exception e)
        {
            ouUnder = "no consensus data";
            ouOver = "no consensus data";
            atsHome = "no consensus data";
            atsAway ="no consensus data";
        }
        ouOversMap.put(MatchupID, ouOver);
        ouUndersMap.put(MatchupID, ouUnder);
        atsHomesMap.put(MatchupID, atsAway);
        atsAwaysMap.put(MatchupID, atsHome);
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
    public HashMap<String, String> getGameIdentifierMap()
    {
        return gameIdentifierMap;
    }
}


