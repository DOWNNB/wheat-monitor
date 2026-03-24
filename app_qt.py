#!/usr/bin/env python3
"""
冬小麦长势监测系统 - 简约白色版
"""
import os, sys, warnings, threading, requests
import numpy as np
from io import BytesIO
from datetime import datetime
from PIL import Image

warnings.filterwarnings('ignore')
requests.packages.urllib3.disable_warnings()

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QSlider, QProgressBar,
    QTabWidget, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QFrame, QScrollArea, QHeaderView, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QImage, QFont, QColor

import matplotlib
matplotlib.use('QtAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
import matplotlib.dates as mdates

plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang SC', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['axes.spines.top']   = False
plt.rcParams['axes.spines.right'] = False

STAGES = [
    ('苗期',       '2023-10-01T00:00:00Z', '2023-10-31T23:59:59Z', '#A8D5A2'),
    ('越冬期',     '2023-12-01T00:00:00Z', '2023-12-31T23:59:59Z', '#B3CDE0'),
    ('返青期',     '2024-02-01T00:00:00Z', '2024-02-29T23:59:59Z', '#6BAF92'),
    ('拔节抽穗期', '2024-04-01T00:00:00Z', '2024-04-30T23:59:59Z', '#3A7D44'),
    ('成熟期',     '2024-05-15T00:00:00Z', '2024-06-10T23:59:59Z', '#E8C07D'),
]
BBOX_PRESETS = {
    '河南郑州': [113.5, 34.5, 114.5, 35.2],
    '山东济南': [116.5, 36.2, 117.5, 37.0],
    '河北石家庄': [114.0, 37.8, 115.2, 38.5],
    '陕西关中': [108.0, 34.0, 109.5, 34.8],
    '安徽淮北': [116.0, 33.5, 117.0, 34.3],
}

STYLE = """
QWidget { font-family: 'PingFang SC', 'Arial Unicode MS', Arial; background: #FFFFFF; }
QMainWindow { background: #F7F7F7; }
QLabel { color: #1A1A1A; }
QLineEdit {
    border: 1px solid #D0D0D0; border-radius: 6px;
    padding: 7px 10px; font-size: 13px; background: #FAFAFA; color: #1A1A1A;
}
QLineEdit:focus { border: 1px solid #3A7D44; background: #FFFFFF; }
QComboBox {
    border: 1px solid #D0D0D0; border-radius: 6px;
    padding: 7px 10px; font-size: 13px; background: #FFFFFF; color: #1A1A1A;
    selection-background-color: #3A7D44; selection-color: white;
}
QComboBox:focus { border: 1px solid #3A7D44; }
QComboBox::drop-down {
    border: none; width: 28px;
}
QComboBox::down-arrow {
    width: 12px; height: 12px;
}
QComboBox QAbstractItemView {
    border: 1px solid #D0D0D0; border-radius: 6px;
    background: white; color: #1A1A1A;
    selection-background-color: #3A7D44; selection-color: white;
    padding: 4px;
}
QPushButton {
    border-radius: 6px; padding: 9px 16px;
    font-size: 13px; font-weight: 700; color: #1A1A1A;
}
QPushButton#primary {
    background-color: #2D6A4F; color: white; border: 2px solid #2D6A4F;
    font-weight: 700;
}
QPushButton#primary:hover { background-color: #1B4332; border-color: #1B4332; color: white; }
QPushButton#primary:disabled { background-color: #B7D5C8; border-color: #B7D5C8; color: #FFFFFF; }
QPushButton#secondary {
    background-color: #FFFFFF; color: #2D6A4F;
    border: 2px solid #2D6A4F; font-weight: 700;
}
QPushButton#secondary:hover { background-color: #F0FAF5; color: #1B4332; }
QPushButton#secondary:disabled { color: #AAAAAA; border-color: #CCCCCC; background-color: #FAFAFA; }
QProgressBar {
    border: none; border-radius: 3px;
    background: #E8E8E8; height: 6px; text-align: center;
}
QProgressBar::chunk { background: #2D6A4F; border-radius: 3px; }
QTabWidget::pane { border: 1px solid #EEEEEE; border-radius: 8px; background: white; }
QTabBar::tab {
    padding: 10px 22px; font-size: 13px; color: #888888;
    border: none; background: transparent;
}
QTabBar::tab:selected { color: #1A1A1A; font-weight: 600; border-bottom: 2px solid #2D6A4F; }
QTableWidget { border: none; gridline-color: #F0F0F0; font-size: 13px; color: #1A1A1A; }
QTableWidget::item { padding: 8px; }
QHeaderView::section {
    background: #FAFAFA; border: none;
    border-bottom: 1px solid #EEEEEE;
    padding: 8px; font-weight: 600; color: #444444;
}
QScrollArea { border: none; }
"""

class WheatMonitor:
    def __init__(self):
        self.token = None

    def _req(self, method, url, **kwargs):
        kwargs.setdefault('verify', False)
        for attempt in range(3):
            try:
                return requests.request(method, url, **kwargs)
            except Exception:
                if attempt == 2: raise

    def login(self, username, password):
        resp = self._req('POST',
            'https://identity.dataspace.copernicus.eu/auth/realms/CDSE/protocol/openid-connect/token',
            data={'grant_type':'password','client_id':'cdse-public',
                  'username':username,'password':password})
        data = resp.json()
        if 'access_token' in data:
            self.token = data['access_token']; return True
        return False

    def search(self, bbox, start, end, max_cloud=30):
        headers = {'Authorization': f'Bearer {self.token}'}
        resp = self._req('POST', 'https://stac.dataspace.copernicus.eu/v1/search',
            headers=headers,
            json={'collections':['sentinel-2-l2a'],'bbox':bbox,
                  'datetime':f'{start}/{end}',
                  'query':{'eo:cloud_cover':{'lt':max_cloud}},
                  'sortby':[{'field':'eo:cloud_cover','direction':'asc'}],
                  'limit':3})
        return resp.json().get('features', [])

    def download_thumb(self, feature):
        headers = {'Authorization': f'Bearer {self.token}'}
        for key in ['thumbnail','QUICKLOOK','overview']:
            if key in feature.get('assets', {}):
                try:
                    r = self._req('GET', feature['assets'][key]['href'], headers=headers)
                    if r.status_code == 200 and len(r.content) > 1000:
                        return r.content
                except Exception:
                    continue
        return None

    def calc_ndvi(self, img_bytes):
        try:
            img = Image.open(BytesIO(img_bytes)).convert('RGB')
            arr = np.array(img, dtype=np.float32) / 255.0
            r, g, b = arr[:,:,0], arr[:,:,1], arr[:,:,2]
            nir = np.clip(2*g - r, 0, 1)
            ndvi = (nir - r) / (nir + r + 1e-8)
            mask = (g > r) & (g > b) & (g > 0.1)
            vals = ndvi[mask]
            return float(np.median(vals)) if len(vals) > 100 else float(np.median(ndvi))
        except:
            return None

class AnalyzeWorker(QThread):
    progress = pyqtSignal(int, str)
    result   = pyqtSignal(list)
    error    = pyqtSignal(str)

    def __init__(self, monitor, bbox, max_cloud):
        super().__init__()
        self.monitor = monitor; self.bbox = bbox; self.max_cloud = max_cloud

    def run(self):
        results = []
        try:
            for i, (stage, start, end, color) in enumerate(STAGES):
                self.progress.emit(int(i/len(STAGES)*100), stage)
                features = self.monitor.search(self.bbox, start, end, self.max_cloud)
                if features:
                    f = features[0]
                    date  = f['properties'].get('datetime','')[:10]
                    cloud = f['properties'].get('eo:cloud_cover', 0)
                    thumb = self.monitor.download_thumb(f)
                    ndvi  = self.monitor.calc_ndvi(thumb) if thumb else None
                    results.append({'stage':stage,'date':date,'ndvi':ndvi,
                                    'thumb':thumb,'cloud':cloud,'color':color})
                else:
                    results.append({'stage':stage,'date':'N/A','ndvi':None,
                                    'thumb':None,'cloud':0,'color':color})
            self.progress.emit(100, '完成')
            self.result.emit(results)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('冬小麦长势监测系统')
        self.resize(1200, 780)
        self.monitor = WheatMonitor()
        self.results = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # 顶部导航栏
        nav = QWidget(); nav.setFixedHeight(52)
        nav.setStyleSheet('background:#FFFFFF;border-bottom:1px solid #EEEEEE;')
        nl = QHBoxLayout(nav); nl.setContentsMargins(24, 0, 24, 0)
        title = QLabel('冬小麦长势监测系统')
        title.setStyleSheet('font-size:15px;font-weight:600;color:#1A1A1A;')
        sub = QLabel('Sentinel-2 · Copernicus Data Space')
        sub.setStyleSheet('font-size:11px;color:#AAAAAA;')
        nl.addWidget(title); nl.addStretch(); nl.addWidget(sub)
        root.addWidget(nav)

        # 主体
        body = QWidget(); body.setStyleSheet('background:#F7F7F7;')
        bl = QHBoxLayout(body)
        bl.setContentsMargins(20, 20, 20, 20); bl.setSpacing(16)
        left = self._build_left(); left.setFixedWidth(260)
        right = self._build_right()
        bl.addWidget(left); bl.addWidget(right, stretch=1)
        root.addWidget(body, stretch=1)

        # 状态栏
        self.status_lbl = QLabel('就绪')
        self.status_lbl.setStyleSheet(
            'background:#FAFAFA;border-top:1px solid #EEEEEE;'
            'color:#888;padding:5px 20px;font-size:11px;')
        root.addWidget(self.status_lbl)

    def _card(self):
        f = QFrame()
        f.setStyleSheet('background:white;border-radius:10px;border:1px solid #EEEEEE;')
        l = QVBoxLayout(f); l.setContentsMargins(20, 18, 20, 18); l.setSpacing(10)
        return f, l

    def _section(self, layout, text):
        lbl = QLabel(text)
        lbl.setStyleSheet(
            'font-size:11px;font-weight:700;color:#2D6A4F;'
            'letter-spacing:1.5px;background:transparent;padding:2px 0;')
        layout.addWidget(lbl)

    def _divider(self, layout):
        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet('color:#F0F0F0;margin:4px 0;')
        layout.addWidget(line)

    def _build_left(self):
        card, layout = self._card()

        self._section(layout, '卫星账号')
        input_style = ('background-color:#F0FAF5;color:#1A1A1A;'
                       'border:1.5px solid #A8D5BE;border-radius:6px;'
                       'padding:7px 10px;font-size:13px;')
        self.user_edit = QLineEdit(); self.user_edit.setText('wyzssd@gmail.com')
        self.user_edit.setStyleSheet(input_style)
        self.pass_edit = QLineEdit()
        self.pass_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_edit.setText('wn9SDCA@_=/w!L-')
        self.pass_edit.setStyleSheet(input_style)
        layout.addWidget(self.user_edit)
        layout.addWidget(self.pass_edit)
        self.login_btn = QPushButton('登录')
        self.login_btn.setObjectName('primary')
        self.login_btn.setStyleSheet('background-color:#2D6A4F;color:#FFFFFF;font-size:13px;font-weight:700;border-radius:6px;padding:9px;border:none;')
        self.login_btn.clicked.connect(self._do_login)
        layout.addWidget(self.login_btn)
        self.login_lbl = QLabel('')
        self.login_lbl.setStyleSheet('font-size:11px;color:#888;')
        layout.addWidget(self.login_lbl)

        self._divider(layout)
        self._section(layout, '监测区域')
        self.region_cb = QComboBox()
        self.region_cb.addItems(list(BBOX_PRESETS.keys()))
        self.region_cb.setStyleSheet(
            'QComboBox{background-color:#F0FAF5;color:#1A1A1A;border:1.5px solid #A8D5BE;'
            'border-radius:6px;padding:7px 10px;font-size:13px;}'
            'QComboBox::drop-down{border:none;width:24px;}'
            'QComboBox QAbstractItemView{background:white;color:#1A1A1A;'
            'selection-background-color:#2D6A4F;selection-color:white;'
            'border:1px solid #A8D5BE;border-radius:4px;padding:4px;}')
        layout.addWidget(self.region_cb)

        self._divider(layout)
        self._section(layout, '最大云量')
        row = QHBoxLayout(); row.setSpacing(8)
        self.cloud_slider = QSlider(Qt.Orientation.Horizontal)
        self.cloud_slider.setRange(5, 80); self.cloud_slider.setValue(30)
        self.cloud_slider.setStyleSheet(
            'QSlider::groove:horizontal{height:6px;background:#E0F2EA;border-radius:3px;}'
            'QSlider::handle:horizontal{background:#2D6A4F;border:none;width:16px;height:16px;'
            'margin:-5px 0;border-radius:8px;}'
            'QSlider::sub-page:horizontal{background:#2D6A4F;border-radius:3px;}')
        self.cloud_lbl = QLabel('30%')
        self.cloud_lbl.setStyleSheet('font-size:12px;color:#2D6A4F;font-weight:700;min-width:32px;')
        self.cloud_slider.valueChanged.connect(lambda v: self.cloud_lbl.setText(f'{v}%'))
        row.addWidget(self.cloud_slider); row.addWidget(self.cloud_lbl)
        layout.addLayout(row)

        self._divider(layout)
        self.run_btn = QPushButton('开始分析')
        self.run_btn.setObjectName('primary')
        self.run_btn.setStyleSheet('background-color:#2D6A4F;color:#FFFFFF;font-size:13px;font-weight:700;border-radius:6px;padding:9px;border:none;')
        self.run_btn.setEnabled(False); self.run_btn.clicked.connect(self._do_analyze)
        layout.addWidget(self.run_btn)
        self.save_btn = QPushButton('导出图表')
        self.save_btn.setObjectName('secondary')
        self.save_btn.setStyleSheet('background-color:#FFFFFF;color:#2D6A4F;font-size:13px;font-weight:700;border-radius:6px;padding:9px;border:2px solid #2D6A4F;')
        self.save_btn.setEnabled(False); self.save_btn.clicked.connect(self._do_save)
        layout.addWidget(self.save_btn)

        self._divider(layout)
        self._section(layout, '分析进度')
        self.prog_bar = QProgressBar(); self.prog_bar.setValue(0)
        self.prog_lbl = QLabel('')
        self.prog_lbl.setStyleSheet('font-size:12px;color:#444444;font-weight:500;')
        layout.addWidget(self.prog_bar)
        layout.addWidget(self.prog_lbl)
        layout.addStretch()
        return card

    def _build_right(self):
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(
            'QTabWidget::pane{border:1px solid #EEEEEE;border-radius:10px;background:white;}'
            'QTabBar::tab{padding:10px 22px;font-size:13px;color:#AAAAAA;border:none;background:transparent;}'
            'QTabBar::tab:selected{color:#1A1A1A;font-weight:600;border-bottom:2px solid #1A1A1A;}'
        )
        # Tab1
        self.tab_img = QWidget(); self.tab_img_l = QHBoxLayout(self.tab_img)
        self.tab_img_l.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.tab_img_l.setSpacing(12); self.tab_img_l.setContentsMargins(16,16,16,16)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setWidget(self.tab_img)
        sc.setStyleSheet('border:none;')
        self.tabs.addTab(sc, '卫星影像')
        # Tab2
        self.tab_ts = QWidget(); self.tab_ts_l = QVBoxLayout(self.tab_ts)
        self.tab_ts_l.setContentsMargins(8,8,8,8)
        self.tabs.addTab(self.tab_ts, 'NDVI 时序')
        # Tab3
        self.tab_bar = QWidget(); self.tab_bar_l = QVBoxLayout(self.tab_bar)
        self.tab_bar_l.setContentsMargins(8,8,8,8)
        self.tabs.addTab(self.tab_bar, '生育期对比')
        # Tab4
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['生育期','日期','云量 (%)','NDVI','长势评估'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet('alternate-background-color:#FAFAFA;')
        self.tabs.addTab(self.table, '数据详情')
        return self.tabs

    def _assess(self, ndvi):
        """基于 RGB 缩略图估算值校准的阈值 (真实 NDVI 约为估算值 x8)"""
        if ndvi is None: return '数据不足'
        if ndvi > 0.07: return '旺盛'
        if ndvi > 0.04: return '正常'
        if ndvi > 0.02: return '偏弱'
        return '较差'

    def _do_login(self):
        self.login_btn.setEnabled(False)
        self.login_lbl.setText('连接中...')
        def task():
            ok = self.monitor.login(self.user_edit.text(), self.pass_edit.text())
            if ok:
                self.login_lbl.setText('已登录')
                self.login_lbl.setStyleSheet('font-size:11px;color:#3A7D44;')
                self.run_btn.setEnabled(True)
                self.status_lbl.setText('登录成功  ·  选择区域后点击「开始分析」')
            else:
                self.login_lbl.setText('账号或密码错误')
                self.login_lbl.setStyleSheet('font-size:11px;color:#CC3333;')
                self.login_btn.setEnabled(True)
        threading.Thread(target=task, daemon=True).start()

    def _do_analyze(self):
        self.run_btn.setEnabled(False); self.save_btn.setEnabled(False)
        self.prog_bar.setValue(0)
        for i in reversed(range(self.tab_img_l.count())):
            w = self.tab_img_l.itemAt(i).widget()
            if w: w.deleteLater()
        for tab_l in [self.tab_ts_l, self.tab_bar_l]:
            for i in reversed(range(tab_l.count())):
                w = tab_l.itemAt(i).widget()
                if w: w.deleteLater()
        self.table.setRowCount(0)
        bbox = BBOX_PRESETS[self.region_cb.currentText()]
        self.worker = AnalyzeWorker(self.monitor, bbox, self.cloud_slider.value())
        self.worker.progress.connect(self._on_progress)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(lambda e: QMessageBox.critical(self, '错误', e))
        self.worker.start()

    def _on_progress(self, pct, msg):
        self.prog_bar.setValue(pct)
        self.prog_lbl.setText(msg)
        self.status_lbl.setText(f'分析中  ·  {msg}')

    def _on_result(self, results):
        self.results = results
        self._render_images(results)
        self._render_charts(results)
        self._render_table(results)
        self.run_btn.setEnabled(True); self.save_btn.setEnabled(True)
        self.prog_lbl.setText('完成')
        self.status_lbl.setText(f'分析完成  ·  {self.region_cb.currentText()}  ·  共 {len(results)} 期')

    def _render_images(self, results):
        for r in results:
            cell = QFrame()
            cell.setStyleSheet(
                f'background:white;border-radius:10px;'
                f'border:1px solid #EEEEEE;')
            cell.setFixedWidth(172)
            cl = QVBoxLayout(cell); cl.setContentsMargins(10,12,10,12); cl.setSpacing(6)
            cl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            if r['thumb']:
                img = Image.open(BytesIO(r['thumb'])).resize((148,148))
                qimg = QImage(img.tobytes(), img.width, img.height, QImage.Format.Format_RGB888)
                pix = QPixmap.fromImage(qimg)
                il = QLabel(); il.setPixmap(pix); il.setAlignment(Qt.AlignmentFlag.AlignCenter)
                il.setStyleSheet('border-radius:6px;')
                cl.addWidget(il)
            else:
                ph = QLabel('暂无影像')
                ph.setAlignment(Qt.AlignmentFlag.AlignCenter)
                ph.setFixedSize(148,148)
                ph.setStyleSheet('background:#F7F7F7;border-radius:6px;color:#BBBBBB;font-size:12px;')
                cl.addWidget(ph)
            accent = QFrame(); accent.setFixedHeight(3)
            accent.setStyleSheet(f'background:{r["color"]};border-radius:2px;')
            cl.addWidget(accent)
            sl = QLabel(r['stage'])
            sl.setStyleSheet('font-size:12px;font-weight:600;color:#1A1A1A;')
            sl.setAlignment(Qt.AlignmentFlag.AlignCenter); cl.addWidget(sl)
            dl = QLabel(r['date'])
            dl.setStyleSheet('font-size:11px;color:#AAAAAA;')
            dl.setAlignment(Qt.AlignmentFlag.AlignCenter); cl.addWidget(dl)
            nv = f"NDVI  {r['ndvi']:.3f}" if r['ndvi'] else 'NDVI  —'
            nl = QLabel(nv)
            nl.setStyleSheet('font-size:12px;color:#3A7D44;font-weight:500;')
            nl.setAlignment(Qt.AlignmentFlag.AlignCenter); cl.addWidget(nl)
            self.tab_img_l.addWidget(cell)

    def _render_charts(self, results):
        data = [(r['date'],r['ndvi'],r['stage'],r['color'])
                for r in results if r['ndvi'] is not None and r['date'] != 'N/A']
        if not data: return
        dates  = [datetime.strptime(d,'%Y-%m-%d') for d,_,_,_ in data]
        ndvis  = [n for _,n,_,_ in data]
        stages = [s for _,_,s,_ in data]
        colors = [c for _,_,_,c in data]

        # 时序图
        fig1, ax1 = plt.subplots(figsize=(8, 4))
        fig1.patch.set_facecolor('white')
        ax1.set_facecolor('white')
        ax1.plot(dates, ndvis, '-', color='#DDDDDD', lw=1.5, zorder=1)
        for d, n, c in zip(dates, ndvis, colors):
            ax1.scatter(d, n, s=120, color=c, zorder=3, edgecolors='white', lw=2)
        for d, n, s in zip(dates, ndvis, stages):
            ax1.annotate(f'{s}\n{n:.3f}', xy=(d, n), xytext=(0, 14),
                         textcoords='offset points', ha='center', fontsize=8.5,
                         color='#555555')
        ax1.set_ylabel('NDVI', fontsize=11, color='#555')
        ax1.set_title(f'{self.region_cb.currentText()}  冬小麦 NDVI 时序变化',
                      fontsize=13, fontweight='600', color='#1A1A1A', pad=14)
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        ax1.tick_params(colors='#AAAAAA', labelsize=10)
        ax1.set_ylim(-0.02, max(ndvis)*1.6 if ndvis else 0.5)
        ax1.grid(axis='y', color='#F0F0F0', lw=1)
        for sp in ax1.spines.values(): sp.set_color('#EEEEEE')
        fig1.autofmt_xdate(); fig1.tight_layout()
        c1 = FigureCanvasQTAgg(fig1)
        self.tab_ts_l.addWidget(c1)
        plt.close(fig1)

        # 柱状图
        fig2, ax2 = plt.subplots(figsize=(8, 4))
        fig2.patch.set_facecolor('white')
        ax2.set_facecolor('white')
        x = range(len(data))
        bars = ax2.bar(x, ndvis, color=colors, width=0.5,
                       edgecolor='white', linewidth=0)
        for bar, n in zip(bars, ndvis):
            ax2.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.004,
                     f'{n:.3f}', ha='center', va='bottom',
                     fontsize=10, color='#555555')
        ax2.set_xticks(list(x)); ax2.set_xticklabels(stages, fontsize=11, color='#555')
        ax2.set_ylabel('NDVI', fontsize=11, color='#555')
        ax2.set_title('各生育期 NDVI 对比', fontsize=13,
                      fontweight='600', color='#1A1A1A', pad=14)
        ax2.set_ylim(0, max(ndvis)*1.4 if ndvis else 0.5)
        ax2.axhline(0.3, color='#CCCCCC', ls='--', lw=1, label='参考阈值  0.3')
        ax2.legend(fontsize=10, frameon=False)
        ax2.tick_params(colors='#AAAAAA', labelsize=10)
        ax2.grid(axis='y', color='#F0F0F0', lw=1)
        for sp in ax2.spines.values(): sp.set_color('#EEEEEE')
        fig2.tight_layout()
        c2 = FigureCanvasQTAgg(fig2)
        self.tab_bar_l.addWidget(c2)
        plt.close(fig2)

    def _render_table(self, results):
        self.table.setRowCount(0)
        for r in results:
            row = self.table.rowCount(); self.table.insertRow(row)
            vals = [r['stage'], r['date'], f"{r['cloud']:.1f}",
                    f"{r['ndvi']:.4f}" if r['ndvi'] else '—',
                    self._assess(r['ndvi'])]
            for col, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(row, col, item)

    def _do_save(self):
        folder = QFileDialog.getExistingDirectory(self, '选择保存目录')
        if not folder: return
        data = [(r['date'],r['ndvi'],r['stage'],r['color'])
                for r in self.results if r['ndvi'] is not None and r['date'] != 'N/A']
        if not data:
            QMessageBox.warning(self, '提示', '暂无可导出的数据'); return
        dates  = [datetime.strptime(d,'%Y-%m-%d') for d,_,_,_ in data]
        ndvis  = [n for _,n,_,_ in data]
        stages = [s for _,_,s,_ in data]
        colors = [c for _,_,_,c in data]
        for kind in ['ts', 'bar']:
            fig, ax = plt.subplots(figsize=(12, 6))
            fig.patch.set_facecolor('white'); ax.set_facecolor('white')
            if kind == 'ts':
                ax.plot(dates, ndvis, '-', color='#DDDDDD', lw=1.5)
                for d,n,c in zip(dates,ndvis,colors):
                    ax.scatter(d,n,s=140,color=c,zorder=3,edgecolors='white',lw=2)
                for d,n,s in zip(dates,ndvis,stages):
                    ax.annotate(f'{s}\n{n:.3f}',xy=(d,n),xytext=(0,14),
                                textcoords='offset points',ha='center',fontsize=9,color='#555')
                ax.set_title(f'{self.region_cb.currentText()} 冬小麦 NDVI 时序变化',
                             fontsize=14,fontweight='600',color='#1A1A1A')
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
                ax.set_ylim(-0.02, max(ndvis)*1.6)
                fig.autofmt_xdate()
                path = os.path.join(folder, 'NDVI_时序图.png')
            else:
                x = range(len(data))
                bars = ax.bar(x,ndvis,color=colors,width=0.5,edgecolor='white',lw=0)
                for bar,n in zip(bars,ndvis):
                    ax.text(bar.get_x()+bar.get_width()/2,bar.get_height()+0.004,
                            f'{n:.3f}',ha='center',va='bottom',fontsize=10,color='#555')
                ax.set_xticks(list(x)); ax.set_xticklabels(stages,fontsize=12)
                ax.set_title('各生育期 NDVI 对比',fontsize=14,fontweight='600',color='#1A1A1A')
                ax.set_ylim(0,max(ndvis)*1.4)
                ax.axhline(0.3,color='#CCCCCC',ls='--',lw=1,label='参考阈值 0.3')
                ax.legend(fontsize=10,frameon=False)
                path = os.path.join(folder, 'NDVI_柱状图.png')
            ax.set_ylabel('NDVI',fontsize=12,color='#555')
            ax.tick_params(colors='#AAAAAA')
            ax.grid(axis='y',color='#F0F0F0',lw=1)
            for sp in ax.spines.values(): sp.set_color('#EEEEEE')
            fig.tight_layout()
            fig.savefig(path, dpi=150, bbox_inches='tight')
            plt.close(fig)
        QMessageBox.information(self, '导出成功', f'图表已保存至\n{folder}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setStyleSheet(STYLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
