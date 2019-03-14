package com.hdc.json;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;

import java.util.List;
import java.util.Map;

/**
 * @Description: 基于fastjson封装的json转换工具类
 * @Author: hdc
 * @Date: 2019/3/14 21:31
 */
public class JsonUtils {
    /**
     * json字符串解析为javaBean对象
     * @param str json字符串
     * @param clazz 指定的javaBean对象
     * @param <T>
     * @return
     */
    public static <T> T jsonToBean(String str,Class<T> clazz) {
        return JSON.parseObject(str, clazz);
    }
    /**
     * json字符串解析为List<javaBean>集合
     * @param str  json字符串
     * @param clazz  指定的javaBean对象
     * @param <T>
     * @return
     */
    public static <T> List<T> jsonToList(String str, Class<T> clazz) {
        return JSON.parseArray(str, clazz);
    }

    /**
     * json字符串解析为 Map<String, Object>
     * @param str json字符串
     * @return
     */
    public static Map<String, Object> jsonToMap(String str){
        Map<String, Object> map = JSONObject.parseObject(str);
        return map;
    }

    /**
     * json字符串解析为JSONArray
     * @param str son字符串
     * @return
     */
    public static JSONArray jsonToArray(String str){
        return JSON.parseArray(str);
    }

    /**
     * json字符串解析为List<Map<String, Object>>
     * @param str json字符串
     * @return List<Map<String, Object>>
     */
    public static List<Map<String, Object>> jsonToListMap(String str) {
        return JSON.parseObject(str, new TypeReference<List<Map<String, Object>>>() {
        });
    }

    /**
     * javaBean对象解析为json字符串
     * @param obj  javaBean对象
     * @return
     */
    public static String beanToJson(Object obj){
        return JSON.toJSONString(obj);
    }

    /**
     *  javaBean对象解析为 Map<String, Object>
     * @param obj  javaBean对象
     * @return
     */
    public static Map<String, Object> beanToMap(Object obj){
        Map<String, Object> map = JSON.parseObject(JSON.toJSONString(obj));
        return map;
    }

    /**
     * Map<String, Object>集合解析为javaBean对象
     * @param map Map集合
     * @param clazz javaBean对象
     * @param <T>
     * @return
     */
    public static <T> T mapToBean(Map<String, Object> map, Class<T> clazz){
        String str = JSON.toJSONString(map);
        return JSON.parseObject(str, clazz);
    }

    /**
     * List集合解析为 JSONArray
     * @param list list集合
     * @return
     */
    public static JSONArray listToArray(List list){
        return JSONArray.parseArray(JSON.toJSONString(list));
    }
}

// https://blog.csdn.net/weixin_41622183/article/details/82734592
//https://blog.csdn.net/weixin_43994414/article/details/85543248
//https://blog.csdn.net/u010310183/article/details/51396290
