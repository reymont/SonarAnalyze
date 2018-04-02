package com.cmi.sonar;

import com.cmi.util.PropertyUtil;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.commons.cli.*;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


public class AnalyzeMain {
    private static Log log = LogFactory.getLog(AnalyzeMain.class);

    public static void main(String[] args) throws Exception {
        String service;
        String port;
        String startTime = null;
        String endTime = null;
        //定义
        Options options = new Options();
        options.addOption("?", false, "list help");//false代表不强制有
        options.addOption("h", false, "sonar server");
        options.addOption("p", false, "sonar port");
        options.addOption("s", false, "start date, example: 2018-03-30");
        options.addOption("e", false, "end date, example: 2018-04-02");

        //解析
        CommandLineParser parser = new DefaultParser();
        CommandLine cmd = parser.parse(options, args);

        //查询交互
        if (cmd.hasOption("?") || args.length == 0) {
            String formatstr = "CLI  cli help";
            HelpFormatter hf = new HelpFormatter();
            hf.printHelp(formatstr, "", options, "");
        }

        if (cmd.hasOption("h")) {
            service = cmd.getOptionValue("h");
        } else {
            service = PropertyUtil.getProperty("analyze.service");
        }

        if (cmd.hasOption("p")) {
            port = cmd.getOptionValue("p");
        } else {
            port = PropertyUtil.getProperty("analyze.port");
        }

        if (cmd.hasOption("e")) {
            startTime = cmd.getOptionValue("e");
        } else {
            //获取当前日期
            Date d = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            endTime = sdf.format(d);
        }

        if (cmd.hasOption("s")) {
            endTime = cmd.getOptionValue("s");
        } else {
            DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
            Date dd = df.parse(endTime);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(dd);
            calendar.add(Calendar.DAY_OF_MONTH, -6);
            startTime = df.format(calendar.getTime());
        }

        AnalyzeMain analyzeMain = new AnalyzeMain();
        Map<String, Map<String, String>> bugdateMap = analyzeMain.analyzeData(service, port);
        List<String> projectList = analyzeMain.projectNameList(service, port);
        try {
            new WriteExcelForXSSF().write(projectList, bugdateMap, startTime, endTime);
            log.info("Sonar Analyze Report Export  Successful");
        } catch (ParseException e) {
            log.error("Date Format Error");
            e.printStackTrace();
        }

    }

    /**
     * 通过Get请求获取Sonar中的数据
     *
     */
    private static String httpGet(String path) {
        String line;
        HttpURLConnection connection;
        InputStream content = null;
        BufferedReader in = null;
        try {
            URL url = new URL(path);
            String encoding = Base64.getEncoder().encodeToString("8a8126cf6634c7e8145661a682ff93627b8c4099:".getBytes("UTF-8"));
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setDoOutput(true);
            connection.setRequestProperty("Authorization", "Basic " + encoding);
            content = connection.getInputStream();
            in = new BufferedReader(new InputStreamReader(content));
            if((line = in.readLine()) != null) {
                return line;
            }
        } catch (Exception e) {
            log.error("Get Infomation Error");
        } finally {
            if (content != null) {
                try {
                    content.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return "";
    }

    /**
     * 拼装数据成格式为Map
     * 包含ProjectName、date、bug数量
     */
    private Map<String, Map<String, String>> analyzeData(String service, String port) {

        Map<String, Map<String, String>> analyzeMap = new HashMap<>();
        //获取所有项目信息
        String projectPath = "http://" + service + ":+" + port + "/api/projects/search?ps=500";
        String projectName = httpGet(projectPath);
        JSONObject json = JSONObject.fromObject(projectName);
        JSONArray jsonArray1=JSONArray.fromObject(json.get("components"));
        for (Object object : jsonArray1) {
            Map<String,String>bugDataMap=new HashMap<>();
            JSONObject jsonObject1 = JSONObject.fromObject(object);
            String keyName = (String) jsonObject1.get("key");
            String name=(String) jsonObject1.get("name");
            String bugDataPath = httpGet("http://" + service + ":" + port + "/api/measures/search_history?metrics=bugs&component=" + keyName);
            JSONObject jsonObject = JSONObject.fromObject(bugDataPath);
            JSONArray jsonArray = JSONArray.fromObject(jsonObject.get("measures"));
            for (Object obj : jsonArray) {
                JSONObject jsonObject2 = JSONObject.fromObject(obj);
                JSONArray jsonArray2 = JSONArray.fromObject(jsonObject2.get("history"));

                for (Object obj2 : jsonArray2) {
                    JSONObject jsonObject3 = JSONObject.fromObject(obj2);
                    String key = (String) jsonObject3.get("date");
                    String value = (String) jsonObject3.get("value");
                    bugDataMap.put(key.substring(0, 10), value);
                }
            }
            analyzeMap.put(name, bugDataMap);
        }
        log.info("analyzeData get Successful"+analyzeMap);
        return analyzeMap;
    }



    private List<String> projectNameList(String service, String port) {
        List<String> projectList = new ArrayList<>();
        String projectPath = "http://" + service + ":" + port + "/api/projects/search?ps=500";
        String projectName = httpGet(projectPath);
        JSONObject json = JSONObject.fromObject(projectName);
        JSONArray jsonArray = JSONArray.fromObject(json.get("components"));
        for (Object obj : jsonArray) {
            JSONObject jsonObject1 = JSONObject.fromObject(obj);
            String name = (String) jsonObject1.get("name");
            projectList.add(name);
        }
        log.info("Project List Get Successful");
        return projectList;
    }

}

