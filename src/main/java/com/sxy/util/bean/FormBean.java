package com.sxy.util.bean;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * @author weitao1
 * @time 2020/02/17
 */
@Getter
@Setter
@ToString
public class FormBean {

    /**
     * 自动评论人数
     */
    private Double automaticReview;

    /**
     * 回访人数
     */
    private Double returnVisit;

    /**
     * 回复人数
     */
    private Double reply;

    /**
     * 该活动总条数
     */
    private Double cnt;

    public FormBean() {
    }

    public FormBean(Double automaticReview, Double returnVisit, Double reply, Double cnt) {
        this.automaticReview = automaticReview;
        this.returnVisit = returnVisit;
        this.reply = reply;
        this.cnt = cnt;
    }

    public void add(Double automaticReview, Double returnVisit, Double reply) {
        this.returnVisit += returnVisit;
        this.automaticReview += automaticReview;
        this.reply += reply;
        this.cnt += 1;
    }
}
