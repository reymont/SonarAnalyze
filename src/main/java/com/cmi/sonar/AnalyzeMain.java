package com.cmi.sonar;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.ParseException;
import java.util.*;


public class AnalyzeMain {
    private static Log log = LogFactory.getLog(AnalyzeMain.class);

    public static void main(String[] args) {
        AnalyzeMain analyzeMain = new AnalyzeMain();
        Map<String, Map<String, String>> bugdateMap = analyzeMain.analyzeData();
        List<String>projectList= analyzeMain.projectNameList();
        String startTime=args[0];
        String endTime=args[1];
        try {
            new WriteExcelForXSSF().write(projectList,bugdateMap,startTime,endTime);
            log.info("Sonar Analyze Report Export  Successful");
        } catch (ParseException e) {
            log.error("Date Format Error");
            e.printStackTrace();
        }

    }

    /**
     * 通过Get请求获取Sonar中的数据
     *
     * @param path
     * @return
     */
    public static String httpGet(String path) {
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
            content = (InputStream) connection.getInputStream();
            in = new BufferedReader(new InputStreamReader(content));
            while ((line = in.readLine()) != null) {
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
    public Map<String, Map<String, String>> analyzeData() {

        Map<String, Map<String, String>> analyzeMap = new HashMap<>();
        //获取所有项目信息
        String projectPath = "http://172.20.62.127:9000/api/projects/search?ps=500";
        String projectName = httpGet(projectPath);
        JSONObject json = JSONObject.fromObject(projectName);
        Map<String, List<Map<String, String>>> projectMap = (Map<String, List<Map<String, String>>>) json;
        List<Map<String, String>> dataList = projectMap.get("components");
        for (Map<String, String> map : dataList) {
            Map<String, String> bugDataMap = new HashMap<>();
            String bugDataPath = httpGet("http://172.20.62.127:9000/api/measures/search_history?metrics=bugs&component=" + map.get("key"));
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
            analyzeMap.put(map.get("name"), bugDataMap);
        }
        log.info("analyzeMap:"+analyzeMap);
        return analyzeMap;
    }

    public List<String> projectNameList(){
        List<String> projectList=new ArrayList<>();
        String projectPath = "http://172.20.62.127:9000/api/projects/search?ps=500";
        String projectName = httpGet(projectPath);
        JSONObject json = JSONObject.fromObject(projectName);
        Map<String, List<Map<String, String>>> projectMap = (Map<String, List<Map<String, String>>>) json;
        List<Map<String, String>> dataList = projectMap.get("components");
        for (Map<String, String> map : dataList) {
            projectList.add(map.get("name"));
        }
        log.info("Project List:"+projectList);
        return projectList;
    }

}