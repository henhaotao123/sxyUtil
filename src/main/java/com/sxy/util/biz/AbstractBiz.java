package com.sxy.util.biz;

import org.apache.log4j.Logger;

/**
 * @author weitao1
 * @time 2020/02/17
 * 抽象的基础biz基类,为了便于拓展
 */
public abstract class AbstractBiz {
    /**
     * 日志logger
     */
    protected final Logger catalinaLog = Logger.getLogger(getClass());
}
