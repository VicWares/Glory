package org.wintrisstech;
/*******************************************************************
 * Covers NFL Extraction Tool
 * Copyright 2020 Dan Farris
 * version Glory 220831
 * Selenium composite version
 *******************************************************************/
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.IOException;

import static org.jsoup.Jsoup.connect;
public class WebSiteReader
{
    private static Document dirtyDoc;
    public static Elements readWebsite(String urlToRead) throws IOException
    {
        System.out.println("WSR20...reading website: " + urlToRead);
        dirtyDoc = Jsoup.parse(String.valueOf(connect(urlToRead).get()));
        return dirtyDoc.getAllElements();
    }
}



