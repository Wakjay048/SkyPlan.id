/**
 * SkyPlan - Kalkulator & Perencana Keuangan Pribadi
 * Dilengkapi fitur Pengingat WhatsApp (Pendekatan 1: wa.me), Tes Bot WA,
 * Ucapan Selamat saat Goals Tercapai, Riwayat Setoran, dan Export Excel (.xlsx).
 */

// ==========================================
// 1. STATE & KONFIGURASI DEFAULT
// ==========================================
const STORAGE_KEY = 'skyplan_finance_data_v3';

function getTodayISODate() {
    const now = new Date();
    const offset = now.getTimezoneOffset();
    const localDate = new Date(now.getTime() - (offset * 60 * 1000));
    return localDate.toISOString().split('T')[0];
}

let appState = {
    income: 5000000,
    userWhatsapp: "",        // Nomor WhatsApp pribadi user (misal: 085269911214)
    expensePercent: 30,
    savingPercent: 50,
    emergencyPercent: 20,
    goals: [
        {
            id: 1700000000001,
            title: "Mobil",
            targetAmount: 30000000,
            reminderDay: 25,
            deposits: [
                { id: 1, amount: 2000, date: "2026-11-13" }
            ]
        },
        {
            id: 1700000000002,
            title: "Dana Darurat & Menikah",
            targetAmount: 30000000,
            reminderDay: 25,
            deposits: [
                { id: 2, amount: 4500000, date: "2026-11-13" }
            ]
        }
    ]
};

let allocationChartInstance = null;

// ==========================================
// 2. UTILITAS FORMAT RUPIAH, TANGGAL & NOMOR WA
// ==========================================

function formatRupiah(number) {
    if (isNaN(number) || number === null) return 'Rp 0';
    return 'Rp ' + Math.round(number).toLocaleString('id-ID');
}

function formatRibuan(number) {
    if (!number && number !== 0) return '';
    return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function parseRibuan(str) {
    if (!str) return 0;
    const cleaned = str.toString().replace(/[^0-9]/g, '');
    return parseInt(cleaned, 10) || 0;
}

/**
 * Mengubah nomor HP lokal (08xx) menjadi format internasional WhatsApp (628xx)
 */
function formatPhoneForWA(rawPhone) {
    if (!rawPhone) return "";
    let digits = rawPhone.toString().replace(/[^0-9]/g, '');
    if (digits.startsWith('0')) {
        digits = '62' + digits.slice(1);
    } else if (digits.startsWith('8')) {
        digits = '62' + digits;
    }
    return digits;
}

function formatExcelDate(isoDate) {
    if (!isoDate) return '-';
    const parts = isoDate.split('-');
    if (parts.length !== 3) return isoDate;
    const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const year = parts[0].slice(-2);
    const monthIndex = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    return `${day}-${months[monthIndex] || parts[1]}-${year}`;
}

function formatTanggalID(isoDate) {
    if (!isoDate) return '-';
    const parts = isoDate.split('-');
    if (parts.length !== 3) return isoDate;
    const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const year = parts[0];
    const monthIndex = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    return `${day} ${months[monthIndex] || parts[1]} ${year}`;
}

function getGoalTotalSaved(goal) {
    if (!goal.deposits || !Array.isArray(goal.deposits)) {
        return goal.initialAmount || 0;
    }
    return goal.deposits.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
}

function attachCurrencyFormatter(inputElement, onUpdateCallback) {
    inputElement.addEventListener('input', (e) => {
        const rawValue = parseRibuan(e.target.value);
        e.target.value = rawValue > 0 ? formatRibuan(rawValue) : '';
        if (onUpdateCallback) onUpdateCallback(rawValue);
    });
}

// ==========================================
// 3. MANAJEMEN PENYIMPANAN (LOCALSTORAGE)
// ==========================================

function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
}

function loadState() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            appState = { ...appState, ...parsed };

            appState.goals = appState.goals.map(g => {
                if (!Array.isArray(g.deposits)) {
                    const initial = Number(g.initialAmount) || 0;
                    return {
                        ...g,
                        deposits: initial > 0 ? [{ id: Date.now(), amount: initial, date: getTodayISODate() }] : []
                    };
                }
                return g;
            });
        } catch (err) {
            console.error("Gagal memuat data dari LocalStorage:", err);
        }
    }
}

// ==========================================
// 4. INISIALISASI DOM & EVENT LISTENERS
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    populateReminderDays();
    loadState();
    syncInputsFromState();
    initChart();
    updateCalculations();
    setupEventListeners();
    checkTodayReminders();
});

function populateReminderDays() {
    const select = document.getElementById('goalReminderDay');
    for (let i = 1; i <= 31; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = `Setiap Tgl ${i}`;
        if (i === 25) option.selected = true;
        select.appendChild(option);
    }
}

function syncInputsFromState() {
    document.getElementById('incomeInput').value = appState.income > 0 ? formatRibuan(appState.income) : '';
    document.getElementById('userWhatsappInput').value = appState.userWhatsapp || '';
    
    document.getElementById('expenseSlider').value = appState.expensePercent;
    document.getElementById('expensePercentInput').value = appState.expensePercent;
    
    document.getElementById('savingSlider').value = appState.savingPercent;
    document.getElementById('savingPercentInput').value = appState.savingPercent;
    
    document.getElementById('emergencySlider').value = appState.emergencyPercent;
    document.getElementById('emergencyPercentInput').value = appState.emergencyPercent;
}

function setupEventListeners() {
    // Input Gaji Bulanan
    attachCurrencyFormatter(document.getElementById('incomeInput'), (val) => {
        appState.income = val;
        saveState();
        updateCalculations();
    });

    // Input Nomor WhatsApp Pribadi
    const waInput = document.getElementById('userWhatsappInput');
    waInput.addEventListener('input', (e) => {
        appState.userWhatsapp = e.target.value.trim();
        saveState();
    });

    // Tombol Uji Coba BOT WA ke Nomor Pribadi
    document.getElementById('btnTestWhatsapp').addEventListener('click', testWhatsAppBotToPersonal);

    // Formatter nominal pada form tambah goal & modal update saldo
    attachCurrencyFormatter(document.getElementById('goalTargetAmount'));
    attachCurrencyFormatter(document.getElementById('goalInitialAmount'));
    attachCurrencyFormatter(document.getElementById('depositAmountInput'));

    // Sinkronisasi Slider & Input Persentase
    bindPercentagePair('expenseSlider', 'expensePercentInput', 'expensePercent');
    bindPercentagePair('savingSlider', 'savingPercentInput', 'savingPercent');
    bindPercentagePair('emergencySlider', 'emergencyPercentInput', 'emergencyPercent');

    // Form Tambah Target Tabungan
    document.getElementById('goalForm').addEventListener('submit', (e) => {
        e.preventDefault();
        addNewGoal();
    });

    // Form Tambah Setoran Baru di dalam Modal
    document.getElementById('addDepositForm').addEventListener('submit', (e) => {
        e.preventDefault();
        handleAddDeposit();
    });

    // Tutup modal jika klik area backdrop luar
    document.getElementById('editGoalModal').addEventListener('click', (e) => {
        if (e.target.id === 'editGoalModal') closeEditModal();
    });
    document.getElementById('congratsModal').addEventListener('click', (e) => {
        if (e.target.id === 'congratsModal') closeCongratsModal();
    });

    // Tombol Export Ringkasan & Reset
    document.getElementById('btnExport').addEventListener('click', copySummaryToClipboard);
    document.getElementById('btnReset').addEventListener('click', resetAllData);
}

function bindPercentagePair(sliderId, inputId, stateKey) {
    const slider = document.getElementById(sliderId);
    const input = document.getElementById(inputId);

    slider.addEventListener('input', (e) => {
        const val = parseInt(e.target.value, 10) || 0;
        input.value = val;
        appState[stateKey] = val;
        saveState();
        updateCalculations();
    });

    input.addEventListener('input', (e) => {
        let val = parseInt(e.target.value, 10);
        if (isNaN(val)) val = 0;
        if (val > 100) val = 100;
        if (val < 0) val = 0;
        
        slider.value = val;
        appState[stateKey] = val;
        saveState();
        updateCalculations();
    });
}

// ==========================================
// 5. FITUR BOT WHATSAPP (PENDEKATAN 1: WA.ME)
// ==========================================

/**
 * Menguji pengiriman pesan BOT WA ke nomor pribadi yang diinput user
 */
function testWhatsAppBotToPersonal() {
    const waNumber = formatPhoneForWA(appState.userWhatsapp);
    if (!waNumber) {
        showToast("Masukkan nomor WhatsApp pribadi Anda terlebih dahulu.");
        document.getElementById('userWhatsappInput').focus();
        return;
    }

    const monthlySaving = (appState.income * appState.savingPercent) / 100;
    let activeGoalsText = "";

    if (appState.goals.length > 0) {
        activeGoalsText = "\n*Daftar Target Menabung Aktif:*\n";
        appState.goals.forEach((g, i) => {
            const saved = getGoalTotalSaved(g);
            const isDone = saved >= g.targetAmount;
            activeGoalsText += `${i + 1}. *${g.title}* — ${isDone ? '✅ TERCAPAI!' : `Terkumpul ${formatRupiah(saved)} / ${formatRupiah(g.targetAmount)} (Tgl ${g.reminderDay}/bln)`}\n`;
        });
    } else {
        activeGoalsText = "\n_Belum ada target tabungan yang ditambahkan._\n";
    }

    const message = 
`*SKYPLAN BOT — UJI COBA PENGINGAT TABUNGAN*

Halo! Sistem pengingat menabung WhatsApp Anda telah terhubung.
Alokasi Tabungan Bulanan: *${formatRupiah(monthlySaving)}/bulan*
${activeGoalsText}
Bot akan membantu mengingatkan Anda setiap tanggal setor yang dijadwalkan hingga seluruh impian tercapai!`;

    openWhatsAppLink(waNumber, message);
}

/**
 * Mengirimkan pesan pengingat atau ucapan selamat ke WhatsApp untuk satu goal spesifik
 */
function sendGoalWhatsAppNotification(goalId) {
    const goal = appState.goals.find(g => g.id === goalId);
    if (!goal) return;

    const waNumber = formatPhoneForWA(appState.userWhatsapp);
    if (!waNumber) {
        showToast("Isi nomor WhatsApp pribadi di Bagian 1 untuk mengirim notifikasi.");
        document.getElementById('userWhatsappInput').focus();
        return;
    }

    const totalSaved = getGoalTotalSaved(goal);
    const remaining = Math.max(0, goal.targetAmount - totalSaved);
    const monthlySaving = (appState.income * appState.savingPercent) / 100;
    const isCompleted = totalSaved >= goal.targetAmount;

    let message = "";

    if (isCompleted) {
        // Ucapan Selamat jika Goals sudah selesai 100%
        message = 
`*SELAMAT! TARGET TABUNGAN TERCAPAI!*

Luar biasa! Anda telah berhasil menyelesaikan target menabung:
Tujuan: *${goal.title}*
Total Terkumpul: *${formatRupiah(totalSaved)}*
Target Dana: *${formatRupiah(goal.targetAmount)}*

Semoga pencapaian ini membawa kebahagiaan dan keberkahan. Tetap semangat merencanakan impian berikutnya bersama *SkyPlan.id*! ✨`;
    } else {
        // Pesan Pengingat Rutin selama periode menabung
        const timeline = calculateGoalTimeline(remaining, monthlySaving);
        const reminder = getNextReminderInfo(goal.reminderDay);

        message = 
` *PENGINGAT MENABUNG — SKYPLAN BOT*

Halo! Jangan lupa menyisihkan dana tabungan untuk impian Anda:
 Tujuan: *${goal.title}*
 Jadwal Setor: *Setiap Tanggal ${goal.reminderDay}* (${reminder.text})
 Rekomendasi Setoran: *${formatRupiah(monthlySaving)}/bulan*

 *Progres Saat Ini:*
• Terkumpul: *${formatRupiah(totalSaved)}*
• Target: *${formatRupiah(goal.targetAmount)}*
• Kurang: *${formatRupiah(remaining)}*
• Estimasi Selesai: *${timeline.monthsText}* (${timeline.estimatedDateText})

Sedikit demi sedikit, lama-lama menjadi bukit. Semangat menabung!`;
    }

    openWhatsAppLink(waNumber, message);
}

function openWhatsAppLink(phone, text) {
    const encoded = encodeURIComponent(text);
    const url = `https://wa.me/${phone}?text=${encoded}`;
    window.open(url, '_blank');
}

/**
 * Mengecek apakah hari ini adalah tanggal menabung untuk salah satu goals yang belum selesai
 */
function checkTodayReminders() {
    const todayDay = new Date().getDate();
    const dueGoals = appState.goals.filter(g => {
        const isNotDone = getGoalTotalSaved(g) < g.targetAmount;
        return isNotDone && g.reminderDay === todayDay;
    });

    const banner = document.getElementById('todayReminderBanner');
    if (dueGoals.length > 0) {
        const firstGoal = dueGoals[0];
        document.getElementById('bannerReminderTitle').textContent = `Hari Ini Jadwal Menabung: "${firstGoal.title}"!`;
        document.getElementById('bannerReminderDesc').textContent = dueGoals.length > 1
            ? `Ada ${dueGoals.length} target tabungan yang dijadwalkan setor setiap tanggal ${todayDay}.`
            : `Jangan lupa menyisihkan dana tabungan bulan ini agar target tercapai tepat waktu.`;

        document.getElementById('btnBannerSendWA').onclick = () => sendGoalWhatsAppNotification(firstGoal.id);
        banner.classList.remove('hidden');
        lucide.createIcons();
    } else {
        banner.classList.add('hidden');
    }
}

// ==========================================
// 6. LOGIKA PERHITUNGAN UTAMA & RENDER UI
// ==========================================

function updateCalculations() {
    const { income, expensePercent, savingPercent, emergencyPercent } = appState;

    const expenseAmount = (income * expensePercent) / 100;
    const savingAmount = (income * savingPercent) / 100;
    const emergencyAmount = (income * emergencyPercent) / 100;

    document.getElementById('expenseNominal').textContent = formatRupiah(expenseAmount);
    document.getElementById('savingNominal').textContent = formatRupiah(savingAmount);
    document.getElementById('emergencyNominal').textContent = formatRupiah(emergencyAmount);
    document.getElementById('summaryMonthlySaving').textContent = `${formatRupiah(savingAmount)}/bulan`;

    const dailyLimit = expenseAmount / 30;
    const weeklyLimit = expenseAmount / 4;
    document.getElementById('dailyLimit').textContent = formatRupiah(dailyLimit);
    document.getElementById('weeklyLimit').textContent = formatRupiah(weeklyLimit);

    const totalPercent = expensePercent + savingPercent + emergencyPercent;
    const badgeTotal = document.getElementById('badgeTotalPercent');
    const warningBox = document.getElementById('percentWarning');

    badgeTotal.textContent = `Total: ${totalPercent}%`;
    if (totalPercent === 100) {
        badgeTotal.className = "px-3 py-1 rounded-full text-xs font-bold bg-emerald-100 text-emerald-700";
        warningBox.classList.add('hidden');
    } else {
        badgeTotal.className = "px-3 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800";
        document.getElementById('currentPercentSum').textContent = `${totalPercent}%`;
        warningBox.classList.remove('hidden');
    }

    updateChart(expenseAmount, savingAmount, emergencyAmount);
    renderGoals(savingAmount);
    checkTodayReminders();
}

// ==========================================
// 7. VISUALISASI GRAFIK (CHART.JS)
// ==========================================

function initChart() {
    const ctx = document.getElementById('allocationChart').getContext('2d');
    allocationChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Pengeluaran & Kebutuhan', 'Tabungan Impian', 'Dana Darurat & Hiburan'],
            datasets: [{
                data: [30, 50, 20],
                backgroundColor: ['#0EA5E9', '#06B6D4', '#818CF8'],
                borderColor: '#FFFFFF',
                borderWidth: 3,
                hoverOffset: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '68%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        usePointStyle: true,
                        padding: 16,
                        font: { family: "'Plus Jakarta Sans', sans-serif", size: 11, weight: '600' },
                        color: '#475569'
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw || 0;
                            return ` ${context.label}: ${formatRupiah(value)}`;
                        }
                    }
                }
            }
        }
    });
}

function updateChart(expense, saving, emergency) {
    if (!allocationChartInstance) return;
    if (appState.income === 0) {
        allocationChartInstance.data.datasets[0].data = [
            appState.expensePercent,
            appState.savingPercent,
            appState.emergencyPercent
        ];
    } else {
        allocationChartInstance.data.datasets[0].data = [expense, saving, emergency];
    }
    allocationChartInstance.update();
}

// ==========================================
// 8. MANAJEMEN TARGET MENABUNG (GOALS)
// ==========================================

function addNewGoal() {
    const titleInput = document.getElementById('goalTitle');
    const targetInput = document.getElementById('goalTargetAmount');
    const initialInput = document.getElementById('goalInitialAmount');
    const reminderSelect = document.getElementById('goalReminderDay');

    const title = titleInput.value.trim();
    const targetAmount = parseRibuan(targetInput.value);
    const initialAmount = parseRibuan(initialInput.value);
    const reminderDay = parseInt(reminderSelect.value, 10);

    if (!title || targetAmount <= 0) {
        showToast("Mohon isi judul dan nominal target dengan benar.");
        return;
    }

    const deposits = [];
    if (initialAmount > 0) {
        deposits.push({
            id: Date.now(),
            amount: initialAmount,
            date: getTodayISODate()
        });
    }

    const newGoal = {
        id: Date.now() + 1,
        title,
        targetAmount,
        reminderDay,
        deposits
    };

    appState.goals.unshift(newGoal);
    saveState();

    titleInput.value = '';
    targetInput.value = '';
    initialInput.value = '';

    updateCalculations();

    if (initialAmount >= targetAmount) {
        showCongratsModal(newGoal);
    } else {
        showToast(`Target "${title}" berhasil ditambahkan!`);
    }
}

function deleteGoal(id) {
    if (!confirm("Hapus target tabungan ini?")) return;
    appState.goals = appState.goals.filter(g => g.id !== id);
    saveState();
    updateCalculations();
    showToast("Target tabungan dihapus.");
}

function getNextReminderInfo(reminderDay) {
    const today = new Date();
    const currentDay = today.getDate();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();

    let targetDate = new Date(currentYear, currentMonth, reminderDay);
    if (currentDay > reminderDay) {
        targetDate = new Date(currentYear, currentMonth + 1, reminderDay);
    }

    const diffTime = targetDate - new Date(currentYear, currentMonth, currentDay);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    const monthNames = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const dateString = `${targetDate.getDate()} ${monthNames[targetDate.getMonth()]} ${targetDate.getFullYear()}`;

    if (diffDays === 0) {
        return { text: `Hari ini! (${dateString})`, urgent: true };
    } else if (diffDays <= 3) {
        return { text: `${diffDays} hari lagi (${dateString})`, urgent: true };
    } else {
        return { text: `${diffDays} hari lagi (${dateString})`, urgent: false };
    }
}

function calculateGoalTimeline(remainingAmount, monthlySaving) {
    if (remainingAmount <= 0) {
        return {
            monthsText: "Tercapai! 🎉",
            estimatedDateText: "Target sudah terpenuhi",
            isAchievable: true
        };
    }

    if (monthlySaving <= 0) {
        return {
            monthsText: "Tidak terhitung",
            estimatedDateText: "Atur gaji & % tabungan > 0",
            isAchievable: false
        };
    }

    const monthsNeeded = Math.ceil(remainingAmount / monthlySaving);
    const years = Math.floor(monthsNeeded / 12);
    const months = monthsNeeded % 12;
    
    let durationParts = [];
    if (years > 0) durationParts.push(`${years} Tahun`);
    if (months > 0) durationParts.push(`${months} Bulan`);
    const monthsText = durationParts.join(' ');

    const futureDate = new Date();
    futureDate.setMonth(futureDate.getMonth() + monthsNeeded);
    const fullMonths = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];
    const estimatedDateText = `${fullMonths[futureDate.getMonth()]} ${futureDate.getFullYear()}`;

    return {
        monthsText,
        estimatedDateText,
        isAchievable: true
    };
}

function renderGoals(monthlySaving) {
    const container = document.getElementById('goalsContainer');
    const emptyState = document.getElementById('emptyGoalsState');

    container.innerHTML = '';

    if (appState.goals.length === 0) {
        emptyState.classList.remove('hidden');
        return;
    }

    emptyState.classList.add('hidden');

    appState.goals.forEach(goal => {
        const totalSaved = getGoalTotalSaved(goal);
        const remaining = Math.max(0, goal.targetAmount - totalSaved);
        const progressPercent = Math.min(100, Math.round((totalSaved / goal.targetAmount) * 100));
        const isCompleted = totalSaved >= goal.targetAmount;
        
        const timeline = calculateGoalTimeline(remaining, monthlySaving);
        const reminder = getNextReminderInfo(goal.reminderDay);

        const card = document.createElement('div');
        card.className = `goal-card bg-white rounded-2xl p-6 border ${isCompleted ? 'border-emerald-300 ring-1 ring-emerald-100' : 'border-sky-100'} shadow-sm flex flex-col justify-between space-y-5`;
        
        // Badge status pengingat vs status selesai
        const statusBadge = isCompleted
            ? `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 mb-2">
                    <i data-lucide="award" class="w-3.5 h-3.5"></i>
                    Selamat! Goals Selesai 100%
               </span>`
            : `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-semibold ${reminder.urgent ? 'bg-amber-100 text-amber-800' : 'bg-sky-50 text-sky-700'} mb-2">
                    <i data-lucide="bell" class="w-3 h-3"></i>
                    Pengingat Tgl ${goal.reminderDay}: ${reminder.text}
               </span>`;

        card.innerHTML = `
            <div>
                <!-- Header Kartu -->
                <div class="flex items-start justify-between gap-2 mb-3">
                    <div>
                        ${statusBadge}
                        <h3 class="font-bold text-lg text-slate-800">${escapeHTML(goal.title)}</h3>
                    </div>

                    <!-- Tombol Aksi: Kirim WA, Export Excel, Edit/Update Saldo, dan Hapus -->
                    <div class="flex items-center gap-1 shrink-0">
                        <button onclick="sendGoalWhatsAppNotification(${goal.id})" title="${isCompleted ? 'Kirim Ucapan Selamat ke WhatsApp' : 'Kirim Pengingat Menabung ke WhatsApp'}"
                            class="p-2 text-emerald-600 bg-emerald-50/70 hover:bg-emerald-100 rounded-lg transition-colors cursor-pointer">
                            <i data-lucide="message-circle" class="w-4 h-4"></i>
                        </button>
                        <button onclick="exportGoalToExcel(${goal.id})" title="Download Laporan Excel (.xlsx)"
                            class="p-2 text-teal-600 hover:bg-teal-50 rounded-lg transition-colors cursor-pointer">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
                        </button>
                        <button onclick="openEditModal(${goal.id})" title="Edit / Update Saldo Terbaru"
                            class="p-2 text-sky-600 hover:bg-sky-50 rounded-lg transition-colors cursor-pointer">
                            <i data-lucide="pencil" class="w-4 h-4"></i>
                        </button>
                        <button onclick="deleteGoal(${goal.id})" title="Hapus Target"
                            class="p-2 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </div>
                </div>

                <!-- Nominal Target & Terkumpul -->
                <div class="flex items-baseline justify-between text-sm mb-2">
                    <span class="text-slate-500">Terkumpul: <strong class="text-slate-800">${formatRupiah(totalSaved)}</strong></span>
                    <span class="font-bold ${isCompleted ? 'text-emerald-600' : 'text-sky-600'}">${formatRupiah(goal.targetAmount)}</span>
                </div>

                <!-- Progress Bar -->
                <div class="w-full h-2.5 ${isCompleted ? 'bg-emerald-100' : 'bg-sky-100'} rounded-full overflow-hidden">
                    <div class="progress-fill h-full ${isCompleted ? 'bg-gradient-to-r from-emerald-400 to-teal-500' : 'bg-gradient-to-r from-sky-400 to-sky-600'} rounded-full" style="width: ${progressPercent}%"></div>
                </div>
                <div class="flex justify-between text-[11px] text-slate-400 mt-1.5">
                    <span>Progres ${progressPercent}% (${goal.deposits ? goal.deposits.length : 0}x setor)</span>
                    <span>${isCompleted ? 'Target Lunas!' : `Kurang: ${formatRupiah(remaining)}`}</span>
                </div>
            </div>

            <!-- Kotak Hasil Estimasi Waktu atau Ucapan Selamat -->
            ${isCompleted ? `
            <div class="p-4 rounded-xl bg-emerald-50/80 border border-emerald-200 flex items-center justify-between gap-3">
                <div>
                    <span class="text-xs font-bold text-emerald-800 block">🎉 Selamat! Impian Tercapai</span>
                    <span class="text-[11px] text-emerald-600">Dana ${formatRupiah(totalSaved)} telah terkumpul penuh.</span>
                </div>
                <button onclick="sendGoalWhatsAppNotification(${goal.id})"
                    class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-bold rounded-lg transition-colors cursor-pointer shrink-0">
                    Kirim ke WA
                </button>
            </div>
            ` : `
            <div class="p-4 rounded-xl bg-sky-50/70 border border-sky-100 grid grid-cols-2 gap-3">
                <div>
                    <span class="text-[11px] text-slate-500 flex items-center gap-1">
                        <i data-lucide="clock" class="w-3.5 h-3.5 text-sky-500"></i>
                        Lama Menabung
                    </span>
                    <strong class="text-sm sm:text-base font-extrabold text-slate-800 block mt-0.5">
                        ${timeline.monthsText}
                    </strong>
                </div>
                <div>
                    <span class="text-[11px] text-slate-500 flex items-center gap-1">
                        <i data-lucide="calendar-check" class="w-3.5 h-3.5 text-sky-500"></i>
                        Estimasi Tercapai
                    </span>
                    <strong class="text-sm sm:text-base font-extrabold text-sky-700 block mt-0.5">
                        ${timeline.estimatedDateText}
                    </strong>
                </div>
            </div>
            `}
        `;

        container.appendChild(card);
    });

    lucide.createIcons();
}

// ==========================================
// 9. MODAL EDIT, RIWAYAT SALDO & UCAPAN SELAMAT
// ==========================================

function openEditModal(goalId) {
    const goal = appState.goals.find(g => g.id === goalId);
    if (!goal) return;

    document.getElementById('editGoalId').value = goal.id;
    document.getElementById('modalGoalTitle').textContent = `Update Saldo: ${goal.title}`;
    document.getElementById('depositAmountInput').value = '';
    document.getElementById('depositDateInput').value = getTodayISODate();

    const modalExcelBtn = document.getElementById('modalDownloadExcelBtn');
    modalExcelBtn.onclick = () => exportGoalToExcel(goal.id);

    renderModalDepositHistory(goal);

    const modal = document.getElementById('editGoalModal');
    modal.classList.remove('hidden');
    modal.classList.add('flex');
    lucide.createIcons();
}

function closeEditModal() {
    const modal = document.getElementById('editGoalModal');
    modal.classList.add('hidden');
    modal.classList.remove('flex');
}

function showCongratsModal(goal) {
    const totalSaved = getGoalTotalSaved(goal);
    document.getElementById('congratsTitle').textContent = `Selamat! Target "${goal.title}" Tercapai!`;
    document.getElementById('congratsMessage').textContent = `Anda telah berhasil mengumpulkan ${formatRupiah(totalSaved)} dari target ${formatRupiah(goal.targetAmount)}. Rayakan pencapaian finansial Anda!`;
    
    document.getElementById('btnSendCongratsWA').onclick = () => {
        sendGoalWhatsAppNotification(goal.id);
    };

    const modal = document.getElementById('congratsModal');
    modal.classList.remove('hidden');
    modal.classList.add('flex');
    lucide.createIcons();
}

function closeCongratsModal() {
    const modal = document.getElementById('congratsModal');
    modal.classList.add('hidden');
    modal.classList.remove('flex');
}

function renderModalDepositHistory(goal) {
    const totalSaved = getGoalTotalSaved(goal);
    document.getElementById('modalCurrentTotal').textContent = formatRupiah(totalSaved);
    document.getElementById('modalTargetTotal').textContent = formatRupiah(goal.targetAmount);
    document.getElementById('modalTableTotal').textContent = formatRupiah(totalSaved);

    const tbody = document.getElementById('depositHistoryBody');
    tbody.innerHTML = '';

    if (!goal.deposits || goal.deposits.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="4" class="py-4 px-3 text-center text-slate-400 italic">
                    Belum ada riwayat setoran saldo.
                </td>
            </tr>
        `;
        return;
    }

    goal.deposits.forEach((dep, index) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80";
        tr.innerHTML = `
            <td class="py-2 px-3 font-medium text-slate-500">${index + 1}</td>
            <td class="py-2 px-3 font-bold text-slate-800">${formatRupiah(dep.amount)}</td>
            <td class="py-2 px-3 text-slate-600">${formatTanggalID(dep.date)}</td>
            <td class="py-2 px-3 text-right">
                <button type="button" onclick="deleteDeposit(${goal.id}, ${dep.id})" title="Hapus baris ini"
                    class="p-1 text-slate-400 hover:text-rose-500 rounded transition-colors cursor-pointer">
                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    lucide.createIcons();
}

function handleAddDeposit() {
    const goalId = parseInt(document.getElementById('editGoalId').value, 10);
    const amount = parseRibuan(document.getElementById('depositAmountInput').value);
    const date = document.getElementById('depositDateInput').value;

    if (amount <= 0 || !date) {
        showToast("Masukkan nominal saldo dan tanggal dengan benar.");
        return;
    }

    const goal = appState.goals.find(g => g.id === goalId);
    if (!goal) return;

    const previousSaved = getGoalTotalSaved(goal);

    if (!Array.isArray(goal.deposits)) goal.deposits = [];
    goal.deposits.push({
        id: Date.now(),
        amount,
        date
    });

    const newTotalSaved = getGoalTotalSaved(goal);

    saveState();
    updateCalculations();
    renderModalDepositHistory(goal);

    document.getElementById('depositAmountInput').value = '';

    // Jika baru saja menyentuh atau melewati target 100%, tutup modal edit & tampilkan ucapan selamat
    if (previousSaved < goal.targetAmount && newTotalSaved >= goal.targetAmount) {
        closeEditModal();
        showCongratsModal(goal);
    } else {
        showToast(`Saldo +${formatRupiah(amount)} berhasil dicatat!`);
    }
}

function deleteDeposit(goalId, depositId) {
    const goal = appState.goals.find(g => g.id === goalId);
    if (!goal) return;

    goal.deposits = goal.deposits.filter(d => d.id !== depositId);
    saveState();
    updateCalculations();
    renderModalDepositHistory(goal);
    showToast("Riwayat setoran dihapus.");
}

// ==========================================
// 10. EXPORT EXCEL (.XLSX) SESUAI FORMAT CONTOH
// ==========================================

function exportGoalToExcel(goalId) {
    const goal = appState.goals.find(g => g.id === goalId);
    if (!goal) return;

    if (typeof XLSX === 'undefined') {
        showToast("Library Excel sedang dimuat, silakan coba beberapa detik lagi.");
        return;
    }

    const deposits = goal.deposits || [];
    const totalDepositsCount = Math.max(deposits.length, 1);

    const wsData = [
        [],
        ["", "", "LAPORAN RIWAYAT TABUNGAN TARGET"],
        ["", "", "Nama Tujuan / Goals", goal.title],
        ["", "", "Target Dana", goal.targetAmount],
        [],
        ["", "No", "Tanggal Memasukkan Saldo", "Saldo yang Dimasukkan (Rp)"]
    ];

    const startDataRow = 7; 
    const endDataRow = startDataRow + totalDepositsCount - 1;

    if (deposits.length === 0) {
        wsData.push(["", 1, "-", 0]);
    } else {
        deposits.forEach((dep, idx) => {
            wsData.push([
                "",
                idx + 1,
                formatExcelDate(dep.date),
                Number(dep.amount) || 0
            ]);
        });
    }

    const totalRowIndex = wsData.length + 1;
    wsData.push([
        "",
        "",
        "Total Saldo",
        { t: 'n', v: getGoalTotalSaved(goal), f: `SUM(D${startDataRow}:D${endDataRow})` }
    ]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Format Rupiah
    if (ws['D4']) ws['D4'].z = '"Rp "#,##0';
    for (let r = startDataRow; r <= startDataRow + totalDepositsCount; r++) {
        const cellRef = `D${r}`;
        if (ws[cellRef]) {
            ws[cellRef].z = '"Rp "#,##0';
        }
    }
    const totalCellRef = `D${totalRowIndex}`;
    if (ws[totalCellRef]) {
        ws[totalCellRef].z = '"Rp "#,##0';
    }

    // Styling Header Kuning
    const headerStyle = {
        fill: { fgColor: { rgb: "FFFF00" } },
        font: { bold: true, color: { rgb: "000000" } },
        alignment: { horizontal: "center", vertical: "center" }
    };

    ['B6', 'C6', 'D6'].forEach(cellRef => {
        if (ws[cellRef]) {
            ws[cellRef].s = headerStyle;
        }
    });

    ws['!cols'] = [
        { wch: 4 },
        { wch: 6 },
        { wch: 30 },
        { wch: 28 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Riwayat Tabungan");

    const safeTitle = goal.title.replace(/[^a-zA-Z0-9_-]/g, '_');
    const fileName = `Tabungan_${safeTitle}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showToast(`File "${fileName}" berhasil diunduh!`);
}

// ==========================================
// 11. FITUR SALIN RINGKASAN, RESET & TOAST
// ==========================================

function copySummaryToClipboard() {
    const { income, expensePercent, savingPercent, emergencyPercent, goals } = appState;
    const savingNominal = (income * savingPercent) / 100;

    let text = `📊 *RENCANA KEUANGAN BULANAN (SkyPlan)*\n`;
    text += `Gaji Bulanan: ${formatRupiah(income)}\n`;
    text += `• Pengeluaran (${expensePercent}%): ${formatRupiah((income * expensePercent) / 100)}\n`;
    text += `• Tabungan Impian (${savingPercent}%): ${formatRupiah(savingNominal)}\n`;
    text += `• Dana Darurat (${emergencyPercent}%): ${formatRupiah((income * emergencyPercent) / 100)}\n\n`;

    if (goals.length > 0) {
        text += `🎯 *DAFTAR TARGET TABUNGAN:*\n`;
        goals.forEach((g, idx) => {
            const saved = getGoalTotalSaved(g);
            const rem = Math.max(0, g.targetAmount - saved);
            const timeline = calculateGoalTimeline(rem, savingNominal);
            text += `${idx + 1}. ${g.title} — Terkumpul: ${formatRupiah(saved)} / ${formatRupiah(g.targetAmount)} | Estimasi: ${timeline.monthsText} (${timeline.estimatedDateText})\n`;
        });
    }

    navigator.clipboard.writeText(text).then(() => {
        showToast("Ringkasan rencana keuangan disalin ke clipboard!");
    }).catch(() => {
        showToast("Gagal menyalin teks.");
    });
}

function resetAllData() {
    if (confirm("Apakah Anda yakin ingin mengosongkan semua data perhitungan dan target?")) {
        localStorage.removeItem(STORAGE_KEY);
        appState = {
            income: 0,
            userWhatsapp: "",
            expensePercent: 50,
            savingPercent: 30,
            emergencyPercent: 20,
            goals: []
        };
        syncInputsFromState();
        updateCalculations();
        showToast("Semua data telah direset.");
    }
}

function showToast(message) {
    const toast = document.getElementById('toast');
    const toastMsg = document.getElementById('toastMessage');
    toastMsg.textContent = message;

    toast.classList.remove('translate-y-20', 'opacity-0');
    setTimeout(() => {
        toast.classList.add('translate-y-20', 'opacity-0');
    }, 3000);
}

function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}