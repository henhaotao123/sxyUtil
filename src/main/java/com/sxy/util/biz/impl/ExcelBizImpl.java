package com.sxy.util.biz.impl;

import com.sxy.util.bean.FormBean;
import com.sxy.util.biz.AbstractBiz;
import com.sxy.util.biz.IExcelBiz;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

/**
 * @author weitao1
 * @time 2020/02/17
 */
@Service
public class ExcelBizImpl extends AbstractBiz implements IExcelBiz {

    @Override
    public InputStream avg(InputStream fis) {

        // 设置数字格式
        jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#0");
        jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
        Map<String, FormBean> dataMap = new HashMap<String, FormBean>(600, 1);
        List<String> keys = new ArrayList<>();

        // 获得工作簿对象
        Workbook originalSheet = null;
        try {
            originalSheet = Workbook.getWorkbook(fis);
        } catch (IOException | BiffException e) {
            e.printStackTrace();
        }
        // 获得所有工作表
        Sheet[] sheets = Objects.requireNonNull(originalSheet).getSheets();
        // 遍历工作表
        if (sheets != null && sheets.length > 0) {
            Sheet sheet = sheets[0];
            // 获得行数
            int rows = sheet.getRows();
            // 读取数据
            for (int row = 1; row < rows; row++) {
                int col = 1;
                Cell cell = sheet.getCell(col, row);
                String id = cell.getContents().trim();
                Double automaticReview = Double.valueOf(sheet.getCell(2, row).getContents());
                Double returnVisit = Double.valueOf(sheet.getCell(3, row).getContents());
                Double reply = Double.valueOf(sheet.getCell(5, row).getContents());
                //dataMap 包含这个 id 将本次的值添加进去
                if (dataMap.containsKey(id)) {
                    dataMap.get(id).add(automaticReview, returnVisit, reply);
                } else {
                    //dataMap 不包含这个 id 创建一个form 传入参数,id 作为 key
                    dataMap.put(id, new FormBean(automaticReview, returnVisit, reply, 1d));
                    keys.add(id);
                }
            }

        }
        originalSheet.close();


        // 创建一个工作簿,承载目标工作表
        WritableWorkbook targetSheet;
        // 创建一个工作表,输出目标文件
        WritableSheet sheet;
        ByteArrayOutputStream os = null;
        try {
            os = new ByteArrayOutputStream();
            targetSheet = Workbook.createWorkbook(os);
            sheet = targetSheet.createSheet("sheet1", 0);
            //表头
            sheet.addCell(new Label(0, 0, "计划id"));
            sheet.addCell(new Label(1, 0, "自动评论人数"));
            sheet.addCell(new Label(2, 0, "回访人数"));
            sheet.addCell(new Label(3, 0, "回访率"));
            sheet.addCell(new Label(4, 0, "回复人数"));
            sheet.addCell(new Label(5, 0, "回复率"));

            int size = dataMap.size();
            for (int row = 1; row <= size; row++) {
                String keyId = keys.get(row - 1);
                FormBean bean = dataMap.get(keyId);
                Double cnt = bean.getCnt();
                sheet.addCell(new Label(0, row, keyId));
                sheet.addCell(new jxl.write.Number(1, row, bean.getAutomaticReview() / cnt, wcfN));
                sheet.addCell(new jxl.write.Number(2, row, bean.getReturnVisit() / cnt, wcfN));
                sheet.addCell(new Label(3, row, ""));
                sheet.addCell(new jxl.write.Number(4, row, bean.getReply() / cnt, wcfN));
                sheet.addCell(new Label(5, row, ""));
            }
            targetSheet.write();
            targetSheet.close();
        } catch (WriteException | IOException e) {
            e.printStackTrace();
        }
        byte[] byteArray = os.toByteArray();
        return new ByteArrayInputStream(byteArray);
    }

    public static InputStream randomLetter() throws IOException {
        String string = "妈妈,老师,爸爸,生活,一起,不是,人们,今天,孩子,心里,奶奶,眼睛,学校,原来,爷爷,地方,过去,事情,以后,可能,晚上,里面,才能,知识,故事,多少,比赛,冬天,所有,样子,成绩,后来,以前,童年,问题,日子,活动,公园,作文,旁边,下午,自然,房间,空气,笑容,明天,风景,音乐,岁月,文化,生气,机会,身影,天气,空中,红色,书包,今年,儿子,汽车,女孩,早晨,道路,认识,办法,精彩,中午,星星,礼物,习惯,树木,女儿,友谊,夜晚,意义,家长,耳朵,家人,门口,班级,人间,厨房,风雨,影响,过年,电话,黄色,种子,广场,清晨,根本,故乡,笑脸,水面,思想,伙伴,美景,照片,水果,彩虹,刚才,月光,先生,鲜花,灯光,感情,亲人,语言,爱心,光明,左右,新年,角落,青蛙,电影,行为,阳台,用心,题目,动力,天地,花园,诗人,树林,雨伞,去年,少女,乡村,对手,上午,分别,活力,作用,古代,公主,力气,从前,作品,空间,黑夜,说明,青年,面包,往事,大小,女人,司机,中心,对面,心头,嘴角,家门,书本,雪人,男人,笑话,云朵,早饭,右手,水平,行人,乐园,花草,人才,左手,目的,课文,优点,年代,灰尘,沙子,小说,儿女,明星,难题,本子,水珠,彩色,路灯,把握,房屋,心愿,左边,新闻,早点,市场,雨点,细雨,书房,毛巾,画家,元旦,绿豆,本领,起点,青菜,土豆,总结,礼貌,右边,窗帘,萝卜,深情,楼房,对话,面条,北方,木头,商店,疑问,后果,现场,诗词,光亮,白菜,男子,大道,风格,梦乡,姐妹,毛衣,园丁,两边,大风,乡下,广播,规定,围巾,意见,大方,头脑,老大,成语,专业,背景,大衣,黄豆,高手,选手,叶片,过往,奶油,时空,大气,借口,抹布,画笔,晚会,山羊,拖鞋,手工,手心,手术,明年,火苗,知己,价格,树苗,主意,摇篮,对比,胖子,专家,年级,落日,东风,屋子,创意,报道,下巴,面子,雪山,迷宫,友好,烟雾,西方,姨妈,问号,年轮,大米,居民,以往,美女,点心,收入,爱人,会议,女士,玩笑,闪光,总理,毛线,发言,春色,月牙,商人,窗花,工业,苦瓜,开水,明日,沙包,见识,大豆,牙膏,父子,头像,净土,长度,风云,面皮,战友,队员,区别,情谊,门票,气象,空地,星座,壁虎,化石,乐队,睡衣,天平,天桥,开头,太太,同窗,马车,格子,日光,亲子,空姐,野外,往日,首都,画图,口才,海马,邮局,场地,午睡,校友,带鱼,拉面,电力,皮带,书画,豆沙,快车,法宝,病床,野餐,乡土,快门,口语,红星,护照,扇贝,队列,老年,说唱,知了,鲜果,正午,电车,香油,梅子,大会,水车,主机,恩情,师长,猎狗,电机,太岁,平头,一向,大全,字画,虾仁,平房,牛皮,冷水,心得,进口,蚜虫,公告,谣言,户口,个子,亲朋,秀色,西风,大话,干果,火把,奶妈,虾皮,苗木,友人,门牙,性别,回声,美工,跑鞋,师父,乐理,王爷,外地,扶手,地主,级别,住房,对虾,午时,公子,君王,火光,童子,头巾,布衣,平面,人气,假牙,文本,进度,少爷,早操,才女,贴画,出路,物体,热带,点子,外语,飘带,后门,干花,花车,京城,常年,游船,用品,爷们,高低,虫牙,借款,老伴,凉席,球迷,军火,预告,禾苗,画像,桥头,高人,水井,菜场,量杯,皮尺,往常,客房,下手,拉力,满怀,处分,样品,好心,白面,风沙,沙土,冰霜,前天,新春,前后,绿灯,日夜,全才,手足,大队,生字,礼堂,园地,出入,满分,晚年,干草,做工,加法,名气,火电,空想,亲友,法力,能手,午觉,会场,反光,乡间,耳目,冷暖,中队,林场,本土,球门,恩师,少数,平地,幼虫,车头,神像,才气,男声,同乡,帮手,车把,渡船,工蚁,两旁,年岁,地洞,先后,乡亲,清早,北边,苦水,言行,全体,头目,长短,生人,叶柄,凉水,起因,远近,小队,后方,把子,小看,拐弯,苗头,跟前,连队,午间,住家,先前,桥洞,常客,下级,冷天,本意,热天,投向,死活,日用,救星,名门,半边,忘性,来人,里边,阳面,车照,活口,生父,长话,前边,浓淡,成人,时候,幸福,中国,作业,微笑,路上,之后,未来,校园,弟弟,心灵,当时,叔叔,困难,人民,决定,眼泪,泪水,脚步,如今,桌子,抬头,食物,姑娘,汗水,要求,现实,信心,印象,泥土,耐心,昨天,责任,劳动,运动,文字,椅子,被子,科学,始终,科技,院子,果实,玻璃,四周,兔子,傍晚,狐狸,课堂,环保,星期,终点,黑板,组织,灯笼,艺术,游客,幸运,兄弟,玉米,猴子,肩膀,理由,报纸,导游,中央,条件,脚印,技术,花坛,附近,钢笔,闹钟,体育,表示,作家,不幸,特点,功夫,楼梯,食品,伤口,灾难,眼光,足迹,泉水,榜样,除夕,荣誉,太空,图画,交通,胳膊,日记,卫生,书桌,物品,花丛,电梯,单位,钱包,粉笔,手表,夏季,冲动,教学,花纹,年华,车站,病人,报告,丈夫,银河,小组,山路,好事,台灯,讲话,烟火,商品,喜庆,大叔,面粉,乐曲,子弹,乘客,句子,图书,外国,好人,火灾,产品,拐杖,泥泞,观点,山坡,纸巾,保安,生死,运气,银行,面纱,农夫,红旗,点滴,夸张,指甲,蛋黄,秋季,胡须,宫殿,助手,特长,旗帜,谈话,祖母,香烟,电灯,封面,奖品,字体,科目,体温,学院,重心,员工,弱点,昨日,礼品,积木,石油,波纹,电器,祖父,祖先,音响,学科,钉子,南极,谜语,如意,信封,纸张,倒影,晴空,骗子,留言,春笋,金秋,图形,条纹,集团,天宫,文件,甲鱼,柳条,客车,菜花,温室,公元,背心,彩票,甲虫,野兔,风波,错觉,极点,纱布,山川,指望,福利,首长,岭南,小号,公历,西服,号角,冰冻,信用,陆地,代理,哈欠,口技,差距,波涛,口味,托福,水波,银子,军舰,旗杆,枝叶,尤物,斗志,百灵,功劳,躺椅,山猫,花费,乘法,腰包,华南,中介,铁塔,双方,附件,欠条,边界,指示,省会,托词,气压,拉杆,奖金,水塔,省份,店面,饭桶,读音,航运,泪花,谷物,界限,歌星,棉布,层次,史书,树桩,纹路,骨气,油灯,功课,怪事,斜阳,伤风,技工,叔父,传达,小弟,类别,表面,雨丝,口令,灾害,水灾,校庆,礼节,记号,半圆,洼地,布匹,近日,治安,年份,冻伤,山冈,泉眼,赤膊,忠告,脑筋,底片,名次,直角,轻舟,谈吐,圆心,底层,坛子,世代,灾区,须知,马匹,风浪,盛会,学员,泪珠,群岛,温饱,京戏,兄长,纺车,农户,拐角,妹夫,山岭,筋骨,便衣,忠心,劳工,哈气,口福,交情,稀客,族长,坏事,脑壳,子孙,封口,渔舟,读物,恩人,景气,戏院,出产,池子,伤亡,山腰,坡地,曲目,谢意,外界,末代,文坛,已往,屋宇,末年,队礼,期中,山洼,饥寒,摇手,密电,沿路,饥民,经历,智慧,旅游,模样,钥匙,黄昏,营养,光彩,旅程,嗓子,主席,品德,规矩,特产,误会,记载,职务,指针,顺序,花冠,输赢,矩形,山脉,志愿,鬼怪,面貌,厚度,正经,样板,木料,学识,新意,农舍,通体,行踪,光芒,桂花,错误,表情,蛋糕,滋味,怀抱,客厅,狮子,例外,缺点,基础,眉毛,资料,卧室,教训,橡皮,状态,火焰,根据,小偷,装饰,习俗,妻子,身躯,脾气,话题,梦幻,眼帘,辣椒,土壤,阶梯,尘埃,法律,片刻,鸽子,农历,枕头,栏杆,悬崖,浑身,程度,氛围,扫帚,同事,骆驼,规律,礼仪,缝隙,花瓶,义务,模型,纪律,年糕,柱子,隧道,无辜,火柴,遭遇,锄头,概念,项链,感触,水稻,规范,耻辱,判断,政策,家训,种类,讲究,围墙,钳子,轮廓,制服,蝙蝠,数据,邮箱,豌豆,秩序,皇冠,海螺,地址,风趣,霞光,奖杯,蛟龙,马蹄,罪恶,名称,餐馆,行程,女巫,砖头,酒吧,编辑,木材,天敌,糕点,竹竿,鸵鸟,梳子,橱柜,惯性,旅馆,论坛,树梢,丰碑,脂肪,俗语,类型,脉搏,堡垒,皮囊,鸿雁,柜台,饥荒,县城,行政,模板,起源,误解,桂冠,茴香,姓氏,电缆,楷模,腊肠,格式,幕布,檀香,雾霭,灵性,虚荣,穴位,墨鱼,胶囊,海盐,山寨,傲骨,锅炉,住宅,链条,垒球,命题,毛毯,缺口,彩绘,鸟瞰,恩惠,岳父,韧性,患者,触角,阅历,酒徒,蒸汽,食盐,隐士,章程,百货,螺旋,董事,外貌,白昼,索道,弃婴,企图,基层,脊背,汽水,碑文,捷报,鉴定,罚款,锁链,毛毡,螺钉,间隔,肤色,肝火,家禽,傲气,山涧,卷尺,浩劫,昼夜,喜讯,空隙,魔杖,袖口,车辙,巡警,气概,落款,学府,幕僚,补贴,秘诀,腔调,尘土,著作,宽度,卷烟,丝绒,行列,稿纸,油桐,顽童,伯父,官职,悬案,模范,壮举,堤坝,赖皮,面颊,屋脊,凡尘,牲畜,美誉,隐患,依据,巢穴,汛期,枕木,惯例,情形,版图,绳索,国魂,卷宗,牢笼,雏鸡,眉目,技艺,枕套,煤渣,油腻,酒席,琼浆,锦缎,火候,丈人,古训,茅屋,沟渠,头绪,孬种,差错,残废,厚爱,迹象,仪态,题字,毡房,列国,尘嚣,功绩,堤岸,赏赐,腊味,瓦匠,肺腑,宪章,孔隙,宗庙,时蔬,邻近,穗子,汤药,原著,梯队,屠刀,财主,躯体,珍禽,孤寡,俗称,住址,绪言,汁液,溪涧,头颈,官兵,末梢,榜首,义举,屠户,衣兜,官名,案情,俗名,螺号,蠢事,例题,锅台,村寨,倦容,油垢,野禽,种畜,案犯,边塞,财势,希望,梦想,精神,历史,兴奋,蝴蝶,代表,热闹,全部,情况,气息,玩具,大概,任务,价值,掌声,关系,形状,实验,绳子,敌人,邻居,情绪,记者,品味,设计,台阶,瓶子,粮食,柿子,回味,嘴唇,当初,旅途,扇子,手掌,打扮,启示,师傅,篮子,集合,珍藏,仙子,魔力,辫子,警告,机械,卡片,天赋,邮票,威风,委员,心扉,重量,绅士,保姆,噪音,菠萝,金属,火炬,片段,簸箕,豹子,学问,匕首,包袱,掩护,眼睑,磨坊,镜片,盗贼,绒毛,硝烟,威严,特征,木薯,极端,灰烬,瑰宝,侄子,造化,熔岩,胶卷,性能,调度,前夕,天性,礼炮,神气,遗迹,胸脯,气节,出息,阻力,秉性,地域,手绢,断言,舞姿,戎装,郊外,次序,四肢,祸患,算术,止境,名堂,歧途,典礼,警觉,看守,冰鞋,才干,船舱,敬意,限期,雹子,阔佬,窘况,翅膀,鼻子,医院,微风,医生,荷花,清香,夜空,友情,生机,讲台,队伍,蜡烛,反应,脖子,晚饭,网络,记录,工具,功能,信息,剪刀,资源,一旦,饼干,建议,观众,大雁,壮观,护士,演绎,钻石,念头,墨水,背包,韵味,衬衫,隔壁,领域,孤儿,消息,路线,围裙,裂缝,寓言,移民,费用,交易,战地,榛子,缘故,血型,渠道,起初,琴师,尾巴,侠客,光景,窟窿,滋养,规模,声望,信箱,长途,畜生,窝头,用途,鼻祖,匣子,论证,琴键,街坊,矿物,渔翁,形容,陈设,幅度,旱灾,劝告,哀思,处境,唾沫,实话,匪徒,性命,音韵,记性,收成,睡梦,救护,蔬菜,摊点,魅力,毅力,沧桑,走廊,咖啡,定格,情趣,帐篷,铠甲,门槛,智商,奢望,海报,风度,公寓,广袤,伴侣,终极,学子,提议,热浪,纪实,怨恨,装潢,契机,功勋,乐土,风韵,鄙人,船舶,后盾,屏蔽,共识,正宗,履历,底细,异议,要素,荆条,着落,体魄,盛名,影壁,元勋,殊荣,吉凶,海域,红晕,纠葛,征兆,诏书,官衔,区区,风物,谗言,新潮,纺锤,家眷,好歹,礼数,笑柄,痴想,羽纱,乞丐,包裹,膝盖,圣诞,昙花,细菌,芦荟,码头,混沌,陨石,剪影,砥砺,引擎,豆蔻,夙愿,疟疾,氮气,悖论,资质,贿赂,簪子,信笺,烟瘾,恶霸,典范,佳肴,泪腺,妃嫔,应酬,书橱,衙门,书斋,差使,心窍,薄礼,臂膊,削壁,噩梦,海蜇,疫病,患难,旨意,精微,苦役,沧海,大厦,喷嚏,烙印,霸道,橄榄,飓风,骨骼,柔情,契约,鹧鸪,消耗,妃子,令堂,口角,癖好,令尊,迟暮,鹊桥,铺垫,粉黛,变故,冥冥,恩怨,离愁,荣辱,首级,胸臆,尸骸,怨愤,妆饰,败绩,死讯,蜗牛,屏幕,内涵,傻瓜,归宿,核心,步骤,啤酒,奇葩,煤炭,创伤,预言,昵称,脚趾,热忱,噱头,口碑,爵位,翌日,洋芋,艾蒿,鱼鳔,肖像,果脯,钻床,累赘,三昧,薄暮,硕果,褥子,暴乱,渔家,生命,最后,东西,人类,社会,眼前,结果,下面,名字,面前,同时,早已,手机,最终,作为,此时,教育,目标,家庭,存在,风筝,电视,一路,烟花,速度,说道,菊花,大人,美味,题记,内容,不时,歌声,伯伯,小区,天下,信念,车子,人家,幻想,最近,沙发,粽子,寒假,态度,个人,瀑布,遗憾,步伐,池塘,月饼,山顶,过后,超市,教官,计划,鞋子,地下,单车,筷子,象征,早餐,大街,脸颊,美食,大雨,思绪,脑海,心态,口袋,假期,眼镜,方面,事业,分数,一头,流水,帽子,地毯,母校,公公,厕所,意志,一身,此刻,女子,小雨,意外,袋子,当年,部分,脚下,跳绳,创新,警察,见证,能量,汤圆,金钱,手臂,乖乖,材料,裙子,现代,婆婆,红包,仙女,上下,家务,裤子,经典,地板,基本,魔法,美德,反复,投入,数字,职业,身材,期间,苦难,外套,今日,矛盾,树干,依靠,眼角,诗句,一旁,村庄,覆盖,事实,特色,传承,事件,光阴,真情,凳子,国王,星球,铭记,卷子,将军,规则,以上,外表,景点,长辈,墙壁,年纪,袜子,名人,城堡,自身,狂风,烈日,五彩,冬季,风采,角色,高度,午后,考虑,莲花,暴雨,飞船,餐厅,春秋,含义,糯米,酒店,以下,效果,标志,叛逆,文具,使命,谎言,爪子,七彩,漫画,亭子,口号,基地,读者,奖状,同志,实力,沉淀,校服,情怀,眉头,同行,身高,以来,上天,鹦鹉,形式,扫把,容颜,猎人,食堂,柜子,歌词,顾客,箱子,书柜,热血,鼻梁,假山,亲戚,全球,稻谷,晚餐,面孔,拖把,地铁,氧气,单词,街头,海豚,代价,秋千,疾病,四处,言语,光线,阶段,部队,企鹅,葫芦,思路,子女,伟人,梦境,红薯,笔记,番茄,主角,水草,物理,盲人,豆子,地位,气势,景区,小品,课本,场面,海面,果树,滑梯,岗位,当地,心声,地狱,工资,人体,最初,心血,正义,负担,点评,前途,风铃,宠物,平台,春季,包装,蟋蟀,精华,大赛,卫士,生态,历程,清泉,书架,奶茶,原则,年货,结构,初春,粉丝,真相,喇叭,连接,地理,午餐,冬至,心事,温情,水分,大厅,一线,天色,水流,大战,湿地,雪糕,魔方,请求,机场,下水,歌手,村子,世人,后人,视力,厨师,重点,优势,游子,杂志,聚会,酒窝,轮椅,技能,事故,光环,寝室,机遇,牌子,景观,墙角,数量,线条,耳机,规划,火锅,幼儿,动静,铲子,小车,庭院,金牌,永生,重任,韭菜,全面,店铺,空白,游人,中年,气候,设备,过客,直线,书签,一览,飞天,都市,梯子,赛车,螳螂,体重,热气,极限,文人,花样,底线,制度,病毒,绿洲,夫妻,湖泊,高原,话筒,范围,棋子,茶几,夜景,淑女,中药,当代,白糖,龙舟,镰刀,花边,云端,软件,捐款,饭桌,茶花,哲学,屋檐,名牌,证书,雕像,必然,官场,榴莲,莲子,干部,掌心,印记,地瓜,女王,大巴,橙子,前程,人工,轻声,油菜,午夜,民间,风扇,婚礼,珠子,诗篇,品牌,导演,窗台,椰子,相识,胶水,侦探,光泽,字母,货车,四方,比分,元素,傻子,马桶,雨季,征文,军装,体力,深度,磁铁,礼服,日期,病房,风帆,翻译,本色,棒子,飞碟,老公,投资,定义,状元,名词,墓碑,外号,霸气,黑人,暮色,定位,智力,章鱼,趣味,皇后,民俗,床单,龙眼,奶酪,造就,藤蔓,稀饭,毽子,分子,孕妇,学业,入口,标语,寺庙,恋爱,齿轮,装备,名单,山林,盆栽,游记,沃土,饭盒,瓷砖,婚姻,职位,因素,米粉,盛世,卡车,鲸鱼,七夕,路面,五岳,煤气,名句,门铃,听力,眼界,中文,霹雳,医学,脚底,飞车,低调,美梦,同胞,云烟,锤子,光头,戒指,佛像,绵羊,战火,恋人,金星,水源,深海,斑马,公众,架子,佳人,北极,雪松,汉子,蒜苗,笨蛋,调料,大将,笑颜,路途,客户,文物,指数,沙拉,疯子,商业,墨镜,宝剑,人大,学历,哨兵,栅栏,弹弓,称呼,舞曲,拉链,铜钱,极致,药店,天幕,黑豆,手帕,无常,评语,直径,战马,服饰,摩托,白酒,外卖,夫妇,月夜,风衣,工地,圣人,油条,先锋,机关,社团,印章,游艇,好运,密度,蟒蛇,农田,猛兽,马蜂,护栏,民生,风范,烤鸭,夹子,当中,小名,春分,春意,宝物,神韵,论文,乐事,春卷,私人,录音,白眼,螺丝,猫眼,水星,二战,冰糖,终生,零点,阴阳,头盔,瓷器,莲藕,高烧,橡胶,中指,品种,皇上,墓地,武警,大便,喜剧,踏板,动态,缺陷,纹理,斑点,凉鞋,资格,带子,羞辱,告白,小鬼,衣架,钟楼,时辰,雀斑,作风,花椒,前世,渔夫,光圈,胶带,家规,凡人,积分,奖牌,焰火,手艺,芋头,表格,影视,蝗虫,烫伤,生肖,小子,甜品,全民,集市,部门,保镖,素养,衬衣,黄油,菜谱,羽翼,外围,剧情,古琴,热风,触手,学术,百科,措施,靴子,山药,农庄,浮萍,蜂窝,东汉,坚果,顶点,箭头,内衣,民谣,药品,晚秋,石膏,航班,大餐,运河,海参,疤痕,快艇,自传,草鱼,妄想,古装,咽喉,趋势,露天,同类,军歌,万世,脚心,苍山,脏话,国宝,情歌,皇家,盾牌,猎枪,俘虏,包公,重力,木雕,武士,驼背,渔网,边城,稻子,志向,人事,天意,尖刀,业务,电流,天窗,餐具,土豪,正月,新房,酷刑,鬓角,象牙,大葱,奇闻,后记,终身,周岁,信鸽,香炉,目录,路标,浴缸,西南,后宫,基金,绣球,字谜,派对,热点,邮件,餐饮,黄酒,嫂子,上古,三星,球赛,烙饼,笔筒,男士,贷款,食谱,信纸,墙纸,大名,花环,万事,鲸鲨,野鸡,托盘,棺材,家居,驿站,宽带,恶意,弹性,坡度,知青,西汉,木槿,蜗居,街景,河豚,跳蚤,驾照,芥末,静电,买单,姻缘,标题,韶华,青瓷,上将,壁纸,流量,虎牙,恶战,猎物,简历,投影,菱形,游轮,纸牌,步枪,贪官,法院,色调,气泡,家电,水枪,家产,水银,书屋,妖精,佛珠,强国,扁豆,泡菜,玩偶,教案,炒面,水彩,元帅,少时,侧面,葬礼,树脂,字幕,芥菜,战歌,刑法,面团,热狗,透视,天花,龙骨,射手,蟠桃,团员,蚊帐,咸菜,银耳,性感,战旗,隔阂,火炉,谜底,人品,甘蓝,童装,合影,皇子,眼皮,苗圃,流程,音速,球鞋,江东,按键,碎花,校花,气功,山城,声响,邮政,预感,前台,米酒,票房,蛋清,刷子,迷途,寒食,少儿,英镑,道家,凉粉,裂痕,支架,语法,虫草,融资,乳名,力士,港币,粤菜,会员,皇族,小惠,绿荫,医药,年历,维族,军刀,长龙,魔芋,烟斗,炼狱,气枪,官家,美金,漏洞,山庄,顿号,英才,空难,麦冬,初雪,花木,腰果,职称,珍宝,东宫,神道,军棋,顺治,薏米,磨牙,行楷,酒色,乌贼,末世,磁力,机油,武装,繁体,橘红,外贸,刀鱼,情诗,夹克,妖孽,家宴,签证,脚本,天线,标本,凉菜,春梦,比邻,钻戒,财经,玄关,猛禽,黑头,棋盘,独家,新书,摘要,精品,游民,高校,曲剧,段子,坐标,出纳,同人,财务,战机,淡妆,菜肴,同谋,神医,乙醇,评剧,美餐,飞艇,对号,耳环,篝火,毛孔,怒火,异性,金婚,传闻,公车,密使,地摊,猛士,公关,通报,黑锅,裁缝,素菜,文员,密会,方程,晚娘,军婚,彩蛋,债券,立春,鞭子,岛屿,急诊,须臾,盒饭,钢管,当前,灵气,砒霜,铁丝,现状,比价,海岛,泥人,邮轮,风湿,五脏,轩辕,请柬,家业,蛔虫,阴谋,面筋,天机,素食,肋骨,寄语,大伯,感想,奸臣,青鱼,车库,文档,荒野,上品,酵母,硬汉,旧书,铁甲,尺码,传媒,时差,幻听,手电,花卷,邪教,产业,酒酿,仲夏,化工,录像,炸药,特工,枪支,扶摇,考场,括号,商界,密室,嫂嫂,通告,颈椎,木工,手镯,猿人,礁石,皮革,寡妇,棱角,眼线,福音,体系,预算,国花,宴会,千秋,茶园,教皇,意念,残骸,发糕,家政,耳光,头饰,村民,军情,单身,丝路,笔名,吧台,华东,牌照,料酒,矿石,尼龙,腰围,焦距,弹痕,黑帮,长空,绝唱,杉木,山芋,山区,鬼魂,飞鸿,声母,时事,水粉,户籍,帝都,编导,杏眼,树懒,勋章,豆豉,家伙,战舰,干货,低烧,高汤,鞋帮,杂碎,宗旨,洞房,画室,钱币,刑警,动词,恋歌,喜宴,苦主,斑秃,影评,主食,甜菜,洁癖,门禁,露台,角钢,锦旗,血沉,冷库,黑糖,情缘,园艺,心肝,鸡胸,旗舰,胎记,蚝油,副词,订单,民法,花魁,食客,大料,水烟,魔窟,带宽,工伤,糙米,夜场,钢板,美色,木炭,干冰,房山,账号,明教,水货,公章,综述,虚数,色盲,乌梅,回合,个体,心率,妇科,熟地,诊所,韭黄,亚麻,表语,比重,客栈,泥垢,复数,土鳖,渔具,炒肝,年假,软卧,赤道,内线,雷暴,酚酞,龙宫,狍子,墓道,权臣,收据,清客,同门,胃酸,纯情,刀片,粟米,泼墨,黑道,传真,禁书,学徒,格调,石材,断片,砚台,孔庙,五金,力学,桌布,夜市,秀发,麻油,暮春,至尊,杠杆,手铐,渔船,教程,巾帼,枪战,土司,鼻尖,前任,泥塑,用户,攻略,修女,团子,提纲,校训,吸盘,秋色,初衷,豪门,都城,晨昏,油墨,风华,手印,资讯,法则,插头,厨具,油饼,扳手,秸秆,佳节,唇膏,东门,百里,蛋白,孝子,面料,楼道,引力,冬装,刺刀,天帝,电压,塔楼,接口,偏见,胡桃,小样,绝色,冰雕,灵犀,锡纸,独舞,斜视,帮派,砧板,军营,狂草,音节,宦官,胎儿,睡袋,后院,电瓶,整数,立体,合金,编码,栗色,税务,流言,氢气,钢丝,抄手,死水,金条,水表,佛门,船员,剑眉,材质,麦芒,门牌,哑巴,袍子,春饼,肾脏,草药,备注,现金,盟友,纪录,弹力,冷饮,金箔,转盘,煤油,玉簪,图表,钩针,扇形,家书,眼珠,鱼子,血色,月宫,摇椅,硬卧,圣旨,波长,铜牌,风电,鼠疫,豆油,而立,资金,底牌,日晷,西点,铁道,壁灯,汇编,校徽,巫术,枪械,寒气,姊妹,原油,氯气,抗体,蜜枣,沙袋,画展,色相,医师,贵妃,宾语,备份,状况,蒸笼,前言,风向,昏君,领结,路径,手段,伤疤,夜宵,花期,发蜡,血战,刑房,杂粮,直觉,诗集,效应,喷子,中水,膏药,娘娘,极度,插曲,明灯,蟹黄,角质,禁忌,药物,除法,天仙,版本,主管,容器,蜜饯,写意,心机,山墙,五味,专辑,图钉,鱼鹰,痞子,手雷,里脊,清茶,面馆,百叶,官员,仪表,新手,喜糖,台历,火眼,课题,斗篷,海湾,剪辑,特效,诸葛,伤痛,剧场,喷漆,吊带,人像,面容,牙医,负债,古书,骗局,猪排,粉条,挂钩,余额,酥油,怪异,小脚,整体,坛坛罐罐,非处方药,无机物,动画片,app,前方高能,可爱,玫瑰,樱花,魔法少女,篮球,丑八怪,月亮,打call,尬舞,DUCK不必,眼神,奥力给,爱了,快乐源泉,上头,杠精,溜了,确认过眼神,C位,神仙打架,买它,哇哦,有内味了,男朋友,女朋友,老婆,小哥哥,小姐姐,辅助,打野,海绵宝宝,二次元,flag,绿茶,freestyle,中秋,迷惑,迷幻,世界,盘他,起飞,梗,谐音更,脱口秀,茶艺,生日,撒娇,可可爱爱,草莓,西瓜,芒果,柚子,橘子,语文,微信,玩吧,语音房,狼人杀,上麦,唱歌,淡黄,长裙,头发,健身,游泳,游戏,英雄联盟,王者荣耀,爱情,一样,喜欢,美丽,一定,美好,开心,明白,重要,经常,真正,害怕,痛苦,干净,辛苦,欢乐,进步,亲爱,完美,金黄,聪明,清新,迷人,共同,直接,真实,听说,可怕,飞快,雪白,着急,乐观,主要,鲜艳,冰冷,细心,奇妙,动人,大量,无知,暖和,正常,平淡,落后,刻苦,晴朗,永久,刚好,相对,平和,广大,秀丽,日常,高级,相同,笔直,安定,冰雪聪明,知足,结实,许久,听话,知名,闷热,众多,拥挤,天生,迷你,老实,友爱,原始,可笑,合格,公共,大红,得力,洁净,暗淡,鲜红,桃红,吓人,多余,秀美,繁忙,冰凉,热心,空旷,冷清,公开,冷淡,齐全,草绿,能干,发火,可心,业余,空心,凉快,长远,土黄,和好,合法,明净,过时,低下,不快,低级,用功,少许,忙乱,要紧,少见,非分,怕人,大忙,特别,伟大,伤心,实在,丰富,同样,巨大,优秀,亲切,讨厌,严厉,积极,整齐,团结,勤劳,负责,难受,绝对,及时,华丽,焦急,良好,相互,惊奇,自觉,好久,温和,亲密,饱满,正式,灵活,开朗,清洁,忠诚,活跃,响亮,干脆,丰满,隐约,贫困,长寿,必要,便利,清贫,客气,沉着,被动,坚决,厚道,年迈,稀奇,扎实,深厚,富丽,牢固,绝情,雄壮,贫苦,灰暗,枯黄,稠密,银灰,切身,洒落,慌乱,烁烁,慌忙,毛纺,劳苦,壮美,困苦,好胜,扎手,灿烂,成熟,善良,美妙,忙碌,顽强,愉快,清晰,真诚,充实,火红,孤单,尊敬,高贵,愉悦,饥饿,艰苦,奇异,苍白,出色,博大,完善,敏捷,渊博,联合,柔美,干燥,悲痛,虚弱,委婉,锐利,勤勉,惨白,夺目,苍翠,幼小,乳白,丧气,坚强,勇敢,孤独,自豪,普通,厉害,陌生,无聊,短暂,尊重,模糊,艰难,晶莹,繁华,翠绿,寂静,欣慰,异常,悄然,顽皮,激烈,崭新,柔和,沮丧,湿润,神圣,绚烂,朴实,配合,皎洁,广阔,悲哀,愧疚,气馁,喧闹,浓郁,欣喜,倒霉,残忍,羞愧,缤纷,坚韧,安详,鼓舞,庞大,可恶,均匀,悦耳,崇高,苦恼,奢侈,浑浊,笨拙,快捷,孤寂,广泛,稳重,平坦,坚毅,抱歉,严谨,超脱,郑重,气愤,羞耻,质朴,奥妙,沉稳,平均,喧哗,历练,孤僻,消瘦,精湛,豪放,嚣张,乏味,灰不溜丢,阴阳怪气,网站,呵呵,表情包,王子,小朋友,太阳,巨蟹座,处女座,短视频,awsl,安排,安利,傲娇,别问,有一说一,确实,标题党,cp,断舍离,购物车,呆萌,打脸,diss,得瑟,大佬,打工,逗逼,肥宅,快乐水,春运,非主流,火星文,佛系,腹黑,肥皂,爱豆,咕咕咕,CD,羊毛,今日份,彩虹屁,柯南,天线宝宝,哆啦A梦,宝藏,宝藏女孩,老实人,鼓励,孤立,卖萌,麦兜,桌游,三国杀,梁非凡,小饼干,女装,社交,聊天,辩论,主持,电影票,康康,深夜,吐槽,神,神吐槽,烧脑,青春,成长,长大,工作,社畜,三十,索然无味,双标,清白,沙雕,诗,远方,飞鸟,刷屏,回复,双击,666,躺枪,他来了,静静,雪花,完整,两开花,咸鱼,自由,了解一下,第一天,小仙女,白鼠,不好,猫,吸猫,小奶狗,小狼狗,嘤嘤嘤,嘤嘤怪,吊打,颜值,一望无际,小拳拳,明信片,漂流瓶,设计图,亲亲,抱抱,举高高,雨神,种草,直播,中二病,真香,纸短情长,扎心,老铁,涨知识,祖安,转发,巅峰,人生,凝视,队友,炸鸡,赞,比心,争议,vlog,匿名,非洲,劳力士,抬棺,垃圾分类,高能,carry,有点东西,不至于,蕾丝,猩猩,布偶,猫猫,套路,光,落地成盒,皮皮虾,醉了,循环,异世界,穿越,黑色,键盘,太美,996,社交网络,海草,抬杠,奶盖,平平无奇,唠嗑,扩列,文学,网抑云,bgm,孙子,魔鬼,教我做事,海王,火,无情,low,汉服,洛丽塔,又双叒叕,群体免疫,匆忙,细胞,濒危物种,小宇宙,路人,宝宝,温柔,追星,跌跌撞撞,离别,滴滴,飞,哈利波特,电竞,jojo,时代,冲浪,爱丽丝,谢邀,刚下飞机,德国,男孩,朋克,锦鲤,官宣,皮,热搜,富士山,春天,夏天,秋天,桃花,大学,漫威,灭霸,江湖,鲈鱼,动物,森林,少年,笔记本,电脑,童话,月色,发现,学习,告诉,成为,回来,充满,离开,出现,得到,发生,回忆,失去,能够,放学,忘记,以为,结束,回答,吃饭,完成,下去,谢谢,休息,直到,进行,加油,关心,注意,变化,起床,前进,跟着,不见,回去,到底,生长,进去,向往,到达,思念,做人,热爱,回头,生病,再见,点头,参观,感激,游玩,跑步,进来,爱护,出生,满意,独立,做事,放心,帮忙,开花,同意,迟到,观看,加入,讨论,开车,打扫,想念,报答,洗澡,引起,发烧,上来,播种,复苏,伸手,无力,开口,书写,返回,舍得,沉睡,代替,商量,搬家,动手,净化,见面,逃跑,提前,漂流,忘却,放手,问好,目送,脸红,消灭,踏青,诉说,改正,眨眼,交往,综合,请问,迷路,居住,清洗,退休,自立,跳远,夸奖,开会,回收,拔河,听课,认错,看病,封闭,交友,跳高,推广,照明,认可,发起,干杯,手写,天明,挂念,谦让,追赶,领先,问答,往来,转让,满座,入迷,闭口,猜想,发电,访问,梳头,爱惜,舍身,会心,造句,冷笑,爬行,认同,缩短,送行,识别,报到,赞同,去火,惊吓,升旗,上山,定向,搭车,办事,当真,吃苦,淡忘,点播,无边,立正,进展,吹拂,动心,相知,点火,出走,失明,排练,起跑,布告,吹风,吓唬,开春,开工,灭火,自习,尽兴,起草,种地,念旧,主办,走红,演唱,带头,说笑,礼让,到期,再现,完工,生火,练功,砍伐,照样,后怕,收口,答谢,看中,独唱,找寻,诉苦,割舍,起立,升学,过门,总计,休养,公用,耕种,借用,借助,从军,练笔,发问,就学,吹捧,直立,相连,不许,列队,出力,闭气,烧火,怕生,进出,看开,改动,多心,自许,找事,叫唤,听讲,喊叫,叫好,发亮,过冬,道贺,多嘴,吹灯,告吹,回升,赶车,走人,看齐,哭诉,收手,苦口,护送,堆放,自量,共处,在位,休想,没完,住口,到头,照办,拉平,叫苦,躲让,练队,变口,觉得,应该,帮助,相信,拥有,读书,接着,感恩,显得,保护,珍惜,自信,消失,期待,追求,介绍,比如,整理,出发,产生,批评,伤害,取得,告别,使用,利用,建设,解决,展现,停止,哭泣,死亡,展开,带领,表扬,欢呼,守护,欺负,建立,写作,打破,包含,沉思,摇头,争取,约定,克服,叹息,问候,宣布,尽量,吸收,讲课,打架,发呆,流浪,计算,打针,转动,围绕,分离,植树,流动,冒险,朗读,停车,购买,破碎,守望,厌恶,等于,装扮,敬爱,纷飞,加班,着迷,交谈,消除,运用,交织,背书,补充,保卫,守候,申请,留学,挑选,流传,挂号,打闹,住宿,识字,解答,仍旧,叹气,环绕,倒车,丢失,纺织,看穿,灰心,休假,求助,躲避,计时,失聪,航行,数落,交际,守卫,针对,合并,保密,互助,染指,呼喊,午休,扮演,休学,跃进,杜绝,保温,染色,探望,浮动,结伴,排斥,钻研,提神,开除,该死,介入,轰动,合计,及第,鸣谢,宇航,埋没,拱手,追问,口算,斥责,打倒,示意,唱戏,上当,争夺,治水,牢记,首肯,寻求,立功,飘浮,切记,串门,搭桥,拍打,丢脸,越过,永别,责令,坐牢,抓紧,须要,乘坐,摸底,拥护,冲凉,抱病,摇动,告状,晃荡,倒台,传话,研读,哄骗,放晴,夺取,过夜,探求,夸大,检视,及格,碰杯,评比,掉队,吞吐,鸣叫,牵动,伸张,食用,出丑,接班,共计,捡拾,运送,上浮,除外,查访,欠安,串通,开采,受苦,垂危,助长,夸口,洒泪,出使,养病,使唤,查办,绕道,引路,丰产,且慢,变易,失掉,笔算,扑闪,伸腰,蹦高,斜射,说穿,欺哄,值得,陪伴,奉献,唠叨,照耀,品尝,发芽,震撼,融化,激励,绝望,荡漾,原谅,道歉,舞动,笼罩,喝彩,承担,思索,不屑,跨越,迷恋,逼近,鄙视,建造,反抗,维持,品读,蕴藏,摩挲,轻蔑,接待,训斥,责怪,堵塞,鼓动,搞鬼,映照,沉没,夸耀,躲藏,预示,指派,埋藏,劝阻,探测,洗劫,开凿,盛产,悠荡,欢闹,选择,属于,宽容,毕业,训练,尝试,信任,敬佩,点燃,呵护,关注,解释,摇曳,传递,提供,停留,适合,弥漫,埋怨,罢了,留恋,分析,呈现,沸腾,邀请,衬托,屹立,开启,崇拜,蔓延,宣传,模仿,奔腾,领悟,推荐,崛起,依赖,凝聚,允许,播放,描述,治疗,打扰,继承,奔驰,淘汰,呼啸,渗透,凝望,阻止,启发,耽误,增强,呼吁,发誓,约束,试图,输入,崩溃,飞驰,经营,奔波,浏览,添加,符合,阻挡,怜悯,担忧,酷爱,抑制,任凭,残缺,侵蚀,泛滥,倾斜,启蒙,悔恨,撒谎,完毕,震动,展览,搁浅,烹饪,降落,辞职,谴责,撞车,昏迷,求婚,缭绕,塑造,窥视,收购,斟酌,辟邪,考证,抢购,扭曲,叮咛,推进,冲破,宛转,筛选,摄像,赞赏,侵犯,扎堆,普及,律动,凝结,遍地,崇尚,寄生,宣泄,坠毁,贯彻,暂停,储存,润泽,落伍,报销,绘图,瓦解,捐赠,爽约,磨炼,订正,担任,剥削,益智,忤逆,逆转,盗窃,防盗,宣誓,困扰,景仰,肇事,移植,召开,预备,解救,赏析,追缴,凝神,撑腰,驱逐,修订,值班,徒劳,无瑕,埋单,邮寄,转载,猜拳,踊跃,辞退,贬值,揣测,栽培,普选,隐蔽,爆满,入伍,游历,省亲,驱使,堪称,遗弃,退伍,鉴别,欺凌,救济,屠戮,寄宿,踮脚,义诊,赏识,镇静,疏远,宽慰,拨号,崩塌,斥资,款待,呼应,重叠,缴纳,搁置,勒令,乞求,伐木,增值,闭塞,防备,堆叠,勒索,猜疑,担待,饭碗,拆除,偷窃,普查,罚球,闲聊,遵从,厘定,远眺,攀谈,浸泡,张贴,烹调,拜见,谢绝,演化,灌输,堆砌,销毁,道谢,欺侮,联结,破损,器重,累积,汇合,购置,丈量,畅销,语塞,连载,变质,偷盗,侍候,辞世,辨认,率领,挫败,仰仗,晾晒,尽职,磨灭,断定,憨笑,责难,遭殃,费劲,阅览,耻笑,探寻,捣乱,信奉,改观,驱散,通晓,闲逛,梳妆,容许,商议,睦邻,抹黑,病愈,久仰,侮蔑,接替,责问,值勤,许配,喷洒,玩赏,超额,义演,浸没,修筑,凑近,斥骂,信守,狂吠,搭乘,收拢,增光,诚服,添置,撞见,扭打,避让,违抗,斥退,闲扯,调运,内详,缀合,淹埋,茁长,障蔽,卷逃,继续,懂得,欣赏,参加,破天荒,不在乎,来得及,巴不得,大团圆,冷处理,想当然,为了,对于,等到,自从,按照,由于,依据,通过,作为,除了,关于,鉴于,依照,经由,除去,比及,自打,除开,可着,一从,为着,错非,由打,至于,打从";
        List<String> strings = new ArrayList<>(Arrays.asList(string.split(",")));
        Collections.shuffle(strings);
        // 创建一个工作簿,承载目标工作表
        WritableWorkbook targetSheet;
        // 创建一个工作表,输出目标文件
        WritableSheet sheet;
        ByteArrayOutputStream os = null;
        Iterator<String> iterator = strings.iterator();
        try {
            os = new ByteArrayOutputStream();
            targetSheet = Workbook.createWorkbook(os);
            sheet = targetSheet.createSheet("sheet1", 0);

            for (int row = 0; row < strings.size() / 3; row++) {
                sheet.addCell(new Label(0, row, iterator.next() + "," + iterator.next() + "," + iterator.next()));
            }
            targetSheet.write();
            targetSheet.close();
        } catch (WriteException | IOException e) {
            e.printStackTrace();
        }
//        File file = new File("E:\\ideaWorkspace\\test.xsl");
        os.writeTo(new FileOutputStream("E:\\ideaWorkspace\\test.xls"));
        byte[] byteArray = os.toByteArray();
        return new ByteArrayInputStream(byteArray);
    }

    public static void main(String[] args) throws IOException {
        randomLetter();

    }
}

