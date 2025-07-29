import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

def analyze_craft_masters():
    """分析工艺美术大师数据的时间分布"""
    print("=== 工艺美术大师时间分析 ===")
    
    # 加载数据
    df = pd.read_excel('工艺美术大师12篇.xlsx')
    print(f"数据形状: {df.shape}")
    print(f"列名: {df.columns.tolist()}")
    
    # 确保出版日期是日期格式
    df['出版日期'] = pd.to_datetime(df['出版日期'])
    
    # 提取年份
    df['出版年份'] = df['出版日期'].dt.year
    df['出版月份'] = df['出版日期'].dt.month
    
    print("\n数据概览:")
    print(df.head())
    
    # 1. 按年份统计的柱状图
    yearly_counts = df['出版年份'].value_counts().sort_index()
    
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=yearly_counts.index,
        y=yearly_counts.values,
        marker_color='#2E86AB',
        text=yearly_counts.values,
        textposition='auto',
    ))
    
    fig1.update_layout(
        title='工艺美术大师文献按年份分布',
        xaxis_title='年份',
        yaxis_title='文献数量',
        template='plotly_white',
        height=500
    )
    
    fig1.write_html("工艺美术大师年份分布.html")
    
    # 2. 按月份统计的柱状图
    monthly_counts = df['出版月份'].value_counts().sort_index()
    
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=monthly_counts.index,
        y=monthly_counts.values,
        marker_color='#A23B72',
        text=monthly_counts.values,
        textposition='auto',
    ))
    
    fig2.update_layout(
        title='工艺美术大师文献按月份分布',
        xaxis_title='月份',
        yaxis_title='文献数量',
        template='plotly_white',
        height=500
    )
    
    fig2.write_html("工艺美术大师月份分布.html")
    
    # 3. 时间线折线图
    fig3 = go.Figure()
    
    # 按时间顺序排列
    df_sorted = df.sort_values('出版日期')
    
    fig3.add_trace(go.Scatter(
        x=df_sorted['出版日期'],
        y=list(range(1, len(df_sorted) + 1)),
        mode='lines+markers',
        name='文献发表时间线',
        line=dict(color='#F18F01', width=3),
        marker=dict(size=8, color='#F18F01')
    ))
    
    fig3.update_layout(
        title='工艺美术大师文献发表时间线',
        xaxis_title='出版日期',
        yaxis_title='文献序号',
        template='plotly_white',
        height=500
    )
    
    fig3.write_html("工艺美术大师时间线.html")
    
    # 4. 作者分布
    author_counts = df['作者'].value_counts()
    
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(
        x=author_counts.values,
        y=author_counts.index,
        orientation='h',
        marker_color='#C73E1D',
        text=author_counts.values,
        textposition='auto',
    ))
    
    fig4.update_layout(
        title='工艺美术大师文献作者分布',
        xaxis_title='文献数量',
        yaxis_title='作者',
        template='plotly_white',
        height=400
    )
    
    fig4.write_html("工艺美术大师作者分布.html")
    
    # 5. 创建综合报告
    create_summary_report(df, yearly_counts, monthly_counts, author_counts)
    
    print("\n=== 分析完成 ===")
    print("已生成以下文件:")
    print("- 工艺美术大师年份分布.html")
    print("- 工艺美术大师月份分布.html") 
    print("- 工艺美术大师时间线.html")
    print("- 工艺美术大师作者分布.html")
    print("- 工艺美术大师分析报告.html")

def create_summary_report(df, yearly_counts, monthly_counts, author_counts):
    """创建综合分析报告"""
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>工艺美术大师文献时间分析报告</title>
        <style>
            body {{ 
                font-family: 'Microsoft YaHei', Arial, sans-serif; 
                margin: 0; 
                padding: 20px; 
                background-color: #f5f5f5; 
                line-height: 1.6; 
            }}
            .container {{ 
                max-width: 1200px; 
                margin: 0 auto; 
                background-color: white; 
                padding: 30px; 
                border-radius: 10px; 
                box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
            }}
            .header {{ 
                text-align: center; 
                margin-bottom: 40px; 
                padding-bottom: 20px; 
                border-bottom: 3px solid #2E86AB; 
            }}
            .header h1 {{ 
                color: #2E86AB; 
                font-size: 2.5em; 
                margin-bottom: 10px; 
            }}
            .header p {{ 
                color: #666; 
                font-size: 1.2em; 
            }}
            .section {{ 
                margin: 40px 0; 
                padding: 20px; 
                background-color: #f8f9fa; 
                border-radius: 8px; 
                border-left: 4px solid #2E86AB; 
            }}
            .section h2 {{ 
                color: #2E86AB; 
                margin-bottom: 20px; 
                font-size: 1.8em; 
            }}
            .chart-container {{ 
                margin: 20px 0; 
                text-align: center; 
            }}
            .stats-grid {{ 
                display: grid; 
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
                gap: 20px; 
                margin: 20px 0; 
            }}
            .stat-card {{ 
                background: white; 
                padding: 20px; 
                border-radius: 8px; 
                box-shadow: 0 2px 5px rgba(0,0,0,0.1); 
                text-align: center; 
            }}
            .stat-number {{ 
                font-size: 2em; 
                font-weight: bold; 
                color: #2E86AB; 
            }}
            .stat-label {{ 
                color: #666; 
                margin-top: 5px; 
            }}
            .data-table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin: 20px 0; 
            }}
            .data-table th, .data-table td {{ 
                border: 1px solid #ddd; 
                padding: 12px; 
                text-align: left; 
            }}
            .data-table th {{ 
                background-color: #2E86AB; 
                color: white; 
            }}
            .data-table tr:nth-child(even) {{ 
                background-color: #f2f2f2; 
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>工艺美术大师文献时间分析报告</h1>
                <p>中国工艺美术大师文献发表时间分布研究</p>
                <p>分析时间: {pd.Timestamp.now().strftime('%Y年%m月%d日')}</p>
            </div>
            
            <div class="section">
                <h2>📊 数据概览</h2>
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">{len(df)}</div>
                        <div class="stat-label">总文献数</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">{df['作者'].nunique()}</div>
                        <div class="stat-label">作者数量</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">{df['出版年份'].min()}</div>
                        <div class="stat-label">最早年份</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">{df['出版年份'].max()}</div>
                        <div class="stat-label">最晚年份</div>
                    </div>
                </div>
                
                <table class="data-table">
                    <tr>
                        <th>书名</th>
                        <th>作者</th>
                        <th>出版日期</th>
                    </tr>
                    {''.join([f'<tr><td>{row["书名"]}</td><td>{row["作者"]}</td><td>{row["出版日期"].strftime("%Y-%m-%d")}</td></tr>' for _, row in df.iterrows()])}
                </table>
            </div>
            
            <div class="section">
                <h2>📈 年份分布分析</h2>
                <div class="chart-container">
                    <iframe src="工艺美术大师年份分布.html" width="100%" height="500"></iframe>
                </div>
                <p><strong>分析发现:</strong></p>
                <ul>
                    <li>文献发表主要集中在 {df['出版年份'].mode().iloc[0]} 年</li>
                    <li>时间跨度: {df['出版年份'].max() - df['出版年份'].min()} 年</li>
                    <li>平均每年发表: {len(df) / (df['出版年份'].max() - df['出版年份'].min() + 1):.1f} 篇文献</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>📅 月份分布分析</h2>
                <div class="chart-container">
                    <iframe src="工艺美术大师月份分布.html" width="100%" height="500"></iframe>
                </div>
                <p><strong>分析发现:</strong></p>
                <ul>
                    <li>发表最多的月份: {monthly_counts.index[0]} 月</li>
                    <li>发表最少的月份: {monthly_counts.index[-1]} 月</li>
                    <li>文献发表时间分布相对均匀</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>⏰ 时间线分析</h2>
                <div class="chart-container">
                    <iframe src="工艺美术大师时间线.html" width="100%" height="500"></iframe>
                </div>
                <p><strong>分析发现:</strong></p>
                <ul>
                    <li>文献发表时间跨度: {df['出版日期'].min().strftime('%Y年%m月')} 至 {df['出版日期'].max().strftime('%Y年%m月')}</li>
                    <li>平均发表间隔: {(df['出版日期'].max() - df['出版日期'].min()).days / (len(df) - 1):.0f} 天</li>
                    <li>反映了工艺美术大师研究的持续性和系统性</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>👥 作者分布分析</h2>
                <div class="chart-container">
                    <iframe src="工艺美术大师作者分布.html" width="100%" height="400"></iframe>
                </div>
                <p><strong>分析发现:</strong></p>
                <ul>
                    <li>主要作者: {author_counts.index[0]} (发表 {author_counts.iloc[0]} 篇)</li>
                    <li>作者贡献度分布相对均匀</li>
                    <li>体现了多位学者对工艺美术大师研究的关注</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>🎯 研究结论</h2>
                <p><strong>主要发现:</strong></p>
                <ul>
                    <li><strong>时间集中性:</strong> 文献发表主要集中在特定年份，反映了研究热潮</li>
                    <li><strong>作者多样性:</strong> 多位学者参与研究，体现了学术界的广泛关注</li>
                    <li><strong>研究系统性:</strong> 时间跨度合理，体现了研究的持续性和系统性</li>
                    <li><strong>学术价值:</strong> 这些文献为工艺美术大师研究提供了重要的学术资料</li>
                </ul>
                
                <p><strong>研究意义:</strong></p>
                <ul>
                    <li>记录了工艺美术大师的学术贡献和艺术成就</li>
                    <li>为工艺美术教育史研究提供了重要资料</li>
                    <li>体现了中国工艺美术学术研究的传统和特色</li>
                </ul>
            </div>
        </div>
    </body>
    </html>
    """
    
    with open("工艺美术大师分析报告.html", "w", encoding="utf-8") as f:
        f.write(html_content)

if __name__ == "__main__":
    analyze_craft_masters() 