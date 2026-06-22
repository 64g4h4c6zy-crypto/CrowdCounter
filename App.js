import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  StyleSheet, View, Text, TouchableOpacity, Alert, Dimensions,
  ScrollView, Platform, AppState
} from 'react-native';
import { CameraView, useCameraPermissions } from 'expo-camera';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import ExcelJS from 'exceljs';

const { width: SW, height: SH } = Dimensions.get('window');

// ==================== 常量 ====================
const COUNT_LINE_Y = 0.55;   // 计数线位置（画面55%高度处 = 门框中线）
const YOUNG_THRESHOLD = 0.55; // 年轻人判定阈值

// ==================== 主组件 ====================
export default function App() {
  const [permission, requestPermission] = useCameraPermissions();
  const [countIn, setCountIn] = useState(0);
  const [countOut, setCountOut] = useState(0);
  const [countYoung, setCountYoung] = useState(0);
  const [isCounting, setIsCounting] = useState(false);
  const [hourlyData, setHourlyData] = useState({});
  const [currentHour, setCurrentHour] = useState('');
  const [lastDetected, setLastDetected] = useState(null);
  const [statusText, setStatusText] = useState('按「开始计数」启动');

  const cooldownRef = useRef(false);
  const prevYRef = useRef(null);
  const trackingRef = useRef({});
  const frameCountRef = useRef(0);
  const youngScoreRef = useRef(0);

  // ==================== 计时更新 ====================
  useEffect(() => {
    const timer = setInterval(() => {
      const now = new Date();
      setCurrentHour(now.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }));
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  // ==================== 保存小时数据 ====================
  useEffect(() => {
    if (!isCounting) return;
    const interval = setInterval(() => {
      const h = new Date().getHours();
      const key = `${h}:00-${h + 1}:00`;
      setHourlyData(prev => ({
        ...prev,
        [key]: {
          in: (prev[key]?.in || 0) + countIn - (prev[key]?._snapshotIn || 0),
          out: (prev[key]?.out || 0) + countOut - (prev[key]?._snapshotOut || 0),
          young: (prev[key]?.young || 0) + countYoung - (prev[key]?._snapshotYoung || 0),
          _snapshotIn: countIn,
          _snapshotOut: countOut,
          _snapshotYoung: countYoung,
        }
      }));
    }, 60000); // 每分钟快照一次
    return () => clearInterval(interval);
  }, [isCounting, countIn, countOut, countYoung]);

  // ==================== 虚拟检测：基于运动轨迹简化逻辑 ====================
  // 在 Expo 托管模式下，我们无法直接访问 Core ML / Vision。
  // 此简化方案通过 CameraView 的人脸/物体检测坐标来估算人体位置。
  // 如需更高精度，可 eject 到裸工作流接入原生 YOLO 模型。
  //
  // 当前策略：利用 expo-camera 的 onBarcodeScanned 的 cornerPoints 估算画面活跃度，
  // 或者使用 onFacesDetected 检测人脸位置（仅用于计数，不存储/识别）。
  
  const handleFacesDetected = useCallback(({ faces }) => {
    if (!isCounting) return;
    
    faces.forEach((face) => {
      const faceId = face.faceID;
      const now = Date.now();
      
      // 面部中心Y坐标（0=顶部, 1=底部）
      const centerY = face.bounds.origin.y + face.bounds.size.height / 2;
      
      // 追踪此人的之前位置
      const tracked = trackingRef.current[faceId];
      
      if (tracked) {
        const prevY = tracked.y;
        const prevTime = tracked.time;
        
        // 判定跨越计数线
        if (prevY < COUNT_LINE_Y && centerY >= COUNT_LINE_Y) {
          // 从上往下 → 进入
          if (!cooldownRef.current) {
            setCountIn(c => c + 1);
            cooldownRef.current = true;
            setTimeout(() => { cooldownRef.current = false; }, 800);
            
            // 年轻人估算：根据面部大小/位置推测
            const faceSize = face.bounds.size.width * face.bounds.size.height;
            if (faceSize > 0.015) { // 更大的脸 ≈ 更近/更年轻？
              youngScoreRef.current += 0.6;
            }
            if (youngScoreRef.current > YOUNG_THRESHOLD) {
              setCountYoung(c => c + 1);
              youngScoreRef.current = 0;
            }
            
            setLastDetected('进');
            setStatusText(`检测到进入 → 总计进:${countIn + 1}`);
          }
        } else if (prevY > COUNT_LINE_Y && centerY <= COUNT_LINE_Y) {
          // 从下往上 → 出去
          if (!cooldownRef.current) {
            setCountOut(c => c + 1);
            cooldownRef.current = true;
            setTimeout(() => { cooldownRef.current = false; }, 800);
            setLastDetected('出');
            setStatusText(`检测到离开 → 总计出:${countOut + 1}`);
          }
        }
      }
      
      // 更新追踪数据
      trackingRef.current[faceId] = { y: centerY, time: now };
    });
    
    // 清理过期的追踪数据（超过3秒）
    const now = Date.now();
    Object.keys(trackingRef.current).forEach(id => {
      if (now - trackingRef.current[id].time > 3000) {
        delete trackingRef.current[id];
      }
    });
  }, [isCounting, countIn, countOut]);

  // ==================== 导出 Excel ====================
  const exportExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('客流统计');
      
      // 标题
      sheet.columns = [
        { header: '时段', key: 'hour', width: 16 },
        { header: '进入(人次)', key: 'in', width: 14 },
        { header: '离开(人次)', key: 'out', width: 14 },
        { header: '净流量', key: 'net', width: 12 },
        { header: '年轻人次', key: 'young', width: 14 },
        { header: '年轻占比', key: 'ratio', width: 12 },
      ];
      
      // 表头样式
      const headerRow = sheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } };
      headerRow.alignment = { horizontal: 'center' };
      
      // 数据行
      const hours = Object.keys(hourlyData).sort();
      if (hours.length === 0) {
        hours.push(`${new Date().getHours()}:00-${new Date().getHours() + 1}:00`);
        hourlyData[hours[0]] = { in: countIn, out: countOut, young: countYoung };
      }
      
      let rowNum = 2;
      hours.forEach(h => {
        const d = hourlyData[h];
        const total = (d.in || 0) + (d.out || 0);
        const ratio = total > 0 ? ((d.young || 0) / total * 100).toFixed(1) + '%' : '0%';
        sheet.addRow({
          hour: h,
          in: d.in || 0,
          out: d.out || 0,
          net: (d.in || 0) - (d.out || 0),
          young: d.young || 0,
          ratio,
        });
        rowNum++;
      });
      
      // 合计行
      const totalIn = hours.reduce((s, h) => s + (hourlyData[h].in || 0), 0);
      const totalOut = hours.reduce((s, h) => s + (hourlyData[h].out || 0), 0);
      const totalYoung = hours.reduce((s, h) => s + (hourlyData[h].young || 0), 0);
      const totalAll = totalIn + totalOut;
      
      const sumRow = sheet.addRow({
        hour: '合计',
        in: totalIn,
        out: totalOut,
        net: totalIn - totalOut,
        young: totalYoung,
        ratio: totalAll > 0 ? (totalYoung / totalAll * 100).toFixed(1) + '%' : '0%',
      });
      sumRow.font = { bold: true };
      sumRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      
      // 保存
      const now = new Date();
      const filename = `客流统计_${now.getMonth() + 1}月${now.getDate()}日.xlsx`;
      const filePath = FileSystem.documentDirectory + filename;
      
      const buffer = await workbook.xlsx.writeBuffer();
      const base64 = arrayBufferToBase64(buffer);
      await FileSystem.writeAsStringAsync(filePath, base64, {
        encoding: FileSystem.EncodingType.Base64,
      });
      
      // 分享
      if (await Sharing.isAvailableAsync()) {
        await Sharing.shareAsync(filePath, {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          dialogTitle: '导出客流统计',
        });
      } else {
        Alert.alert('已保存', `文件已保存至: ${filePath}`);
      }
    } catch (e) {
      Alert.alert('导出失败', e.message);
    }
  };

  // ==================== 重置 ====================
  const reset = () => {
    setCountIn(0);
    setCountOut(0);
    setCountYoung(0);
    setHourlyData({});
    setStatusText('已重置');
    trackingRef.current = {};
    youngScoreRef.current = 0;
  };

  // ==================== 权限 ====================
  if (!permission) {
    return <View style={styles.center}><Text style={styles.text}>加载中...</Text></View>;
  }
  if (!permission.granted) {
    return (
      <View style={styles.center}>
        <Text style={styles.text}>需要摄像头权限才能统计客流</Text>
        <TouchableOpacity style={styles.btn} onPress={requestPermission}>
          <Text style={styles.btnText}>授予权限</Text>
        </TouchableOpacity>
      </View>
    );
  }

  // ==================== 主界面 ====================
  return (
    <View style={styles.container}>
      {/* 摄像头 */}
      <CameraView
        style={styles.camera}
        facing="back"
        mode="picture"
        onFacesDetected={isCounting ? handleFacesDetected : undefined}
        faceDetectorSettings={{
          mode: 'fast',
          detectLandmarks: 'none',
          runClassifications: 'none',
          minDetectionInterval: 500,
          tracking: true,
        }}
      >
        {/* 计数线 */}
        <View style={[styles.countLine, { top: `${COUNT_LINE_Y * 100}%` }]}>
          <Text style={styles.lineText}>── 计数线 ──</Text>
        </View>
        
        {/* 状态指示 */}
        {lastDetected && (
          <View style={[
            styles.dot,
            { backgroundColor: lastDetected === '进' ? '#00FF88' : '#FF4488' }
          ]} />
        )}
      </CameraView>

      {/* 数据面板 */}
      <View style={styles.panel}>
        <Text style={styles.statusText}>{statusText}</Text>
        
        <View style={styles.row}>
          <View style={styles.statBox}>
            <Text style={styles.statNum}>{countIn}</Text>
            <Text style={styles.statLabel}>进入</Text>
          </View>
          <View style={styles.statBox}>
            <Text style={styles.statNum}>{countOut}</Text>
            <Text style={styles.statLabel}>离开</Text>
          </View>
          <View style={styles.statBox}>
            <Text style={styles.statNum}>{countIn - countOut}</Text>
            <Text style={styles.statLabel}>净流量</Text>
          </View>
          <View style={styles.statBox}>
            <Text style={[styles.statNum, { color: '#FFD700' }]}>{countYoung}</Text>
            <Text style={styles.statLabel}>年轻人</Text>
          </View>
        </View>

        {/* 按钮 */}
        <View style={styles.btnRow}>
          <TouchableOpacity
            style={[styles.btn, isCounting && styles.btnActive]}
            onPress={() => {
              setIsCounting(!isCounting);
              setStatusText(isCounting ? '已暂停' : '计数中...');
            }}
          >
            <Text style={styles.btnText}>{isCounting ? '⏸ 暂停' : '▶ 开始计数'}</Text>
          </TouchableOpacity>
          
          <TouchableOpacity style={styles.btnExport} onPress={exportExcel}>
            <Text style={styles.btnText}>📊 导出Excel</Text>
          </TouchableOpacity>
          
          <TouchableOpacity style={styles.btnReset} onPress={reset}>
            <Text style={styles.btnText}>↺</Text>
          </TouchableOpacity>
        </View>
      </View>
    </View>
  );
}

// ==================== 工具函数 ====================
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

// ==================== 样式 ====================
const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#000' },
  center: { flex: 1, justifyContent: 'center', alignItems: 'center', backgroundColor: '#0A0A1A', padding: 20 },
  text: { color: '#00F0FF', fontSize: 16, textAlign: 'center' },
  camera: { flex: 1 },
  countLine: {
    position: 'absolute', left: 0, right: 0,
    alignItems: 'center', zIndex: 10,
  },
  lineText: { color: '#00F0FF', fontSize: 14, opacity: 0.7, textShadowColor: '#00F0FF', textShadowRadius: 4 },
  dot: {
    position: 'absolute', top: 12, right: 12,
    width: 12, height: 12, borderRadius: 6, opacity: 0.9,
  },
  panel: {
    backgroundColor: '#111122', padding: 16,
    borderTopWidth: 1, borderTopColor: '#00F0FF33',
  },
  statusText: { color: '#888', fontSize: 12, textAlign: 'center', marginBottom: 10 },
  row: { flexDirection: 'row', justifyContent: 'space-around', marginBottom: 14 },
  statBox: { alignItems: 'center' },
  statNum: { color: '#00F0FF', fontSize: 28, fontWeight: '300' },
  statLabel: { color: '#666', fontSize: 11, marginTop: 2 },
  btnRow: { flexDirection: 'row', justifyContent: 'center', gap: 12 },
  btn: {
    backgroundColor: '#00F0FF22', borderWidth: 1, borderColor: '#00F0FF',
    borderRadius: 20, paddingHorizontal: 24, paddingVertical: 10,
  },
  btnActive: { backgroundColor: '#FF004422', borderColor: '#FF0044' },
  btnExport: {
    backgroundColor: '#00FF8822', borderWidth: 1, borderColor: '#00FF88',
    borderRadius: 20, paddingHorizontal: 20, paddingVertical: 10,
  },
  btnReset: {
    backgroundColor: '#FFFFFF11', borderWidth: 1, borderColor: '#FFFFFF33',
    borderRadius: 20, paddingHorizontal: 14, paddingVertical: 10,
  },
  btnText: { color: '#00F0FF', fontSize: 14, fontWeight: '600' },
});
