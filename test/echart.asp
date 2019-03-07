<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<title>入门</title>
<script src="echarts.js"></script>
<script src="../static/js/sleeplib.js"></script>
</head>
<body>
    <h1>开始测试</h1>
    <hr>
    <!-- 先准备一个用于盛放图表的容器 -->
     <script>
        //通过 echarts.init 方法初始化一个 echarts 实例并通过 setOption 方法生成一个简单的柱状图
        //基于准备好的DOM，实例化echarts实例
        var myChart = echarts.init(document.getElementById("container"));
        var option1 = {
            title : {
                text : 'ECharts 入门案例'
            },
            tooltip : {
                text : '鼠标放上去之后的悬浮提示语句！'
            },
            legend : {
                data : ['销量']
            },
            xAxis : {
                data : ['衬衫','羊毛衫','雪纺衫','裤子','高跟鞋','袜子','内裤']
            },
            yAxis : {},
            series : [ {
                name : '销量',
                type : 'bar',
                data : [ 7, 20, 36, 10, 10, 20, 28 ]
            } ]
        };
        // 使用上面的配置项作为参数，传给echart来显示
        myChart.setOption(option1);
    </script>

why display ng
</body>
</html>
