import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip, Legend } from 'recharts';
import './App.css';

function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await axios.get('http://localhost:3000/api/data');
        setData(response.data);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };
    
    fetchData();
    // Auto refresh every 5 seconds for real-time feel
    const interval = setInterval(fetchData, 5000);
    return () => clearInterval(interval);
  }, []);

  if (loading && !data) return <div className="loading">⏳ Memuat data...</div>;
  if (error) return <div className="error">❌ Error: Pastikan Bot Node.js (API) sedang berjalan!<br/><span style={{fontSize: '1rem'}}>{error}</span></div>;
  if (!data) return <div className="error">❌ Tidak ada data</div>;

  const pieData = [
    { name: 'Pemasukan', value: data.totalPemasukan, color: '#4ade80' },
    { name: 'Pengeluaran', value: data.totalPengeluaran, color: '#f87171' }
  ].filter(item => item.value > 0);

  const formatRp = (num) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR' }).format(num);

  return (
    <div className="dashboard-container">
      <header className="header">
        <h1>WenFinance Dashboard</h1>
        <p>Ringkasan Keuangan - {data.month}</p>
      </header>

      <div className="summary-cards">
        <div className="card glass-panel">
          <h3>Total Pemasukan</h3>
          <p className="amount income">{formatRp(data.totalPemasukan)}</p>
        </div>
        <div className="card glass-panel">
          <h3>Total Pengeluaran</h3>
          <p className="amount expense">{formatRp(data.totalPengeluaran)}</p>
        </div>
        <div className="card glass-panel">
          <h3>Saldo Saat Ini</h3>
          <p className="amount balance">{formatRp(data.saldo)}</p>
        </div>
      </div>

      <div className="charts-section">
        <div className="glass-panel">
          <h2 style={{marginTop: 0, textAlign: 'center'}}>Distribusi Keuangan</h2>
          <div className="chart-container">
            {pieData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={pieData}
                    cx="50%"
                    cy="50%"
                    innerRadius={60}
                    outerRadius={80}
                    paddingAngle={5}
                    dataKey="value"
                  >
                    {pieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <RechartsTooltip formatter={(value) => formatRp(value)} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            ) : (
              <p>Belum ada data untuk grafik.</p>
            )}
          </div>
        </div>

        <div className="glass-panel transactions-list">
          <h2>10 Transaksi Terakhir</h2>
          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Tanggal</th>
                  <th>Keterangan</th>
                  <th>Tipe</th>
                  <th>Jumlah</th>
                </tr>
              </thead>
              <tbody>
                {data.recentTransactions.map((tx, idx) => (
                  <tr key={idx}>
                    <td>{tx.tanggal}</td>
                    <td>{tx.keterangan}</td>
                    <td>
                      <span className={`type-badge ${tx.tipe === 'Pemasukan' ? 'income' : 'expense'}`}>
                        {tx.tipe}
                      </span>
                    </td>
                    <td style={{ fontWeight: 'bold' }}>{formatRp(tx.jumlah)}</td>
                  </tr>
                ))}
                {data.recentTransactions.length === 0 && (
                  <tr>
                    <td colSpan="4" style={{ textAlign: 'center' }}>Belum ada transaksi</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
