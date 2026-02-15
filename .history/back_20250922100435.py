<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🍳 Monitor Cucina v3.0 - Ordini in Tempo Reale</title>
    
    <!-- Firebase SDK -->
    <script src="https://www.gstatic.com/firebasejs/9.23.0/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore-compat.js"></script>
    
    <style>
        /* ===== RESET E BASE ===== */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            min-height: 100vh;
            color: #333;
        }

        /* ===== HEADER ===== */
        .header {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            padding: 20px;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
            margin-bottom: 30px;
        }

        .header h1 {
            font-size: 2.8em;
            color: #1e3c72;
            margin-bottom: 5px;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1);
            font-weight: 700;
        }

        .company-subtitle {
            font-size: 1.2em;
            color: #2a5298;
            font-weight: 600;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 2px;
        }

        /* ===== STATUS BAR ===== */
        .status {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 20px;
            flex-wrap: wrap;
        }

        .status-item {
            display: flex;
            align-items: center;
            gap: 8px;
            background: rgba(46, 204, 113, 0.1);
            padding: 8px 16px;
            border-radius: 25px;
            border: 2px solid #2ecc71;
            transition: all 0.3s ease;
        }

        .status-item.error {
            background: rgba(231, 76, 60, 0.1);
            border-color: #e74c3c;
        }

        .status-dot {
            width: 12px;
            height: 12px;
            background: #2ecc71;
            border-radius: 50%;
            animation: pulse 2s infinite;
        }

        .status-item.error .status-dot {
            background: #e74c3c;
        }

        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
        }

        /* ===== CONTAINER ===== */
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 20px;
        }

        /* ===== NAVIGATION ===== */
        .navigation {
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 15px;
            flex-wrap: wrap;
        }

        .back-button {
            background: linear-gradient(45deg, #3498db, #2980b9);
            color: white;
            border: none;
            padding: 12px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 16px;
            display: none;
            transition: all 0.3s ease;
        }

        .back-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }

        .breadcrumb {
            color: white;
            font-size: 18px;
            font-weight: 600;
        }

        /* ===== FILTERS ===== */
        .filters {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            padding: 20px;
            border-radius: 15px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        }

        .filters-row {
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }

        .filter-label {
            font-weight: 600;
            color: #2c3e50;
            font-size: 14px;
        }

        .filter-input {
            padding: 10px 15px;
            border: 2px solid #bdc3c7;
            border-radius: 8px;
            font-size: 14px;
            transition: border-color 0.3s ease;
        }

        .filter-input:focus {
            outline: none;
            border-color: #3498db;
        }

        .filter-button {
            background: linear-gradient(45deg, #27ae60, #2ecc71);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
        }

        .filter-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(46, 204, 113, 0.3);
        }

        .filter-button.active {
            background: linear-gradient(45deg, #e67e22, #f39c12);
        }

        .filter-button.da-fare {
            background: linear-gradient(45deg, #e74c3c, #c0392b);
        }

        .filter-button.da-fare:hover {
            box-shadow: 0 4px 12px rgba(231, 76, 60, 0.3);
        }

        /* ===== SELECTION BAR ===== */
        .selection-bar {
            margin-top: 15px;
            padding: 15px;
            background: rgba(52, 152, 219, 0.1);
            border: 2px solid #3498db;
            border-radius: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }

        .selection-info {
            font-weight: 600;
            color: #2c3e50;
        }

        .selection-actions {
            display: flex;
            gap: 10px;
        }

        .action-button {
            background: linear-gradient(45deg, #3498db, #2980b9);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
        }

        .action-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }

        .action-button.secondary {
            background: linear-gradient(45deg, #95a5a6, #7f8c8d);
        }

        .action-button.secondary:hover {
            box-shadow: 0 4px 12px rgba(149, 165, 166, 0.3);
        }

        /* ===== ORDERS GRID ===== */
        .orders-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .order-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
            cursor: pointer;
            border: 2px solid transparent;
        }

        .order-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
            border-color: #3498db;
        }

        .order-card.selected {
            border-color: #3498db;
            background: rgba(52, 152, 219, 0.1);
            box-shadow: 0 0 15px rgba(52, 152, 219, 0.3);
        }

        .order-card.selection-mode {
            cursor: pointer;
            position: relative;
        }

        .order-card.selection-mode::before {
            content: '';
            position: absolute;
            top: 10px;
            right: 10px;
            width: 20px;
            height: 20px;
            border: 2px solid #bdc3c7;
            border-radius: 4px;
            background: white;
        }

        .order-card.selected.selection-mode::before {
            background: #3498db;
            border-color: #3498db;
        }

        .order-card.selected.selection-mode::after {
            content: '✓';
            position: absolute;
            top: 12px;
            right: 14px;
            color: white;
            font-size: 12px;
            font-weight: bold;
        }

        .order-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }

        .order-title {
            font-size: 18px;
            font-weight: 700;
            color: #2c3e50;
        }

        .order-badge {
            padding: 4px 8px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            color: white;
        }

        .badge-new {
            background: #e74c3c;
        }

        .badge-modified {
            background: #f39c12;
        }

        .order-info {
            margin-bottom: 15px;
        }

        .info-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
            font-size: 14px;
        }

        .info-label {
            font-weight: 600;
            color: #7f8c8d;
        }

        .info-value {
            color: #2c3e50;
        }

        .order-products {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 10px;
            margin-top: 10px;
        }

        .products-summary {
            font-size: 12px;
            color: #7f8c8d;
            text-align: center;
        }

        /* ===== DETAIL VIEW ===== */
        .detail-view {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
            display: none;
        }

        .detail-header {
            margin-bottom: 30px;
            text-align: center;
        }

        .detail-title {
            font-size: 24px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 10px;
        }

        .detail-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .info-item {
            text-align: center;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 10px;
        }

        .food-section {
            margin-bottom: 30px;
        }

        .section-title {
            font-size: 20px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 15px;
            text-align: center;
        }

        .food-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
            gap: 15px;
        }

        .food-item {
            background: #f8f9fa;
            border-radius: 10px;
            padding: 15px;
            border: 2px solid transparent;
            transition: all 0.3s ease;
        }

        .food-item.completed {
            background: #d5f4e6;
            border-color: #27ae60;
        }

        .food-name {
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 5px;
        }

        .food-quantity {
            color: #7f8c8d;
            font-size: 14px;
        }

        /* ===== FOOTER ===== */
        .footer {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            padding: 15px;
            text-align: center;
            margin-top: 30px;
            border-radius: 15px 15px 0 0;
        }

        .footer p {
            color: #7f8c8d;
            font-size: 14px;
        }

        /* ===== RESPONSIVE ===== */
        @media (max-width: 768px) {
            .header h1 {
                font-size: 2em;
            }
            
            .filters-row {
                flex-direction: column;
                align-items: stretch;
            }
            
            .orders-grid {
                grid-template-columns: 1fr;
            }
            
            .detail-info {
                grid-template-columns: 1fr;
            }
        }

        /* ===== TOTALS VIEW ===== */
        .totals-view {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        }

        .totals-header {
            margin-bottom: 30px;
            text-align: center;
        }

        .totals-title {
            font-size: 24px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 15px;
        }

        .totals-info {
            display: flex;
            justify-content: center;
            gap: 30px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }

        .totals-stat {
            text-align: center;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 10px;
            min-width: 120px;
        }

        .stat-number {
            font-size: 24px;
            font-weight: 700;
            color: #3498db;
        }

        .stat-label {
            font-size: 14px;
            color: #7f8c8d;
            margin-top: 5px;
        }

        .totals-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
        }

        .total-item {
            background: #f8f9fa;
            border-radius: 10px;
            padding: 20px;
            border-left: 5px solid #3498db;
            transition: all 0.3s ease;
        }

        .total-item:hover {
            transform: translateX(5px);
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }

        .total-name {
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 10px;
            font-size: 16px;
        }

        .total-quantity {
            color: #3498db;
            font-size: 24px;
            font-weight: 700;
        }

        .total-orders {
            color: #7f8c8d;
            font-size: 12px;
            margin-top: 5px;
        }

        /* ===== LOADING ===== */
        .loading {
            text-align: center;
            padding: 50px;
            color: white;
            font-size: 18px;
        }

        .spinner {
            border: 4px solid rgba(255, 255, 255, 0.3);
            border-radius: 50%;
            border-top: 4px solid white;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        /* ===== FOOD ITEMS DETTAGLI ===== */
.food-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 15px;
    margin-top: 20px;
}

.food-item {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border-radius: 12px;
    padding: 18px;
    border: 2px solid transparent;
    transition: all 0.3s ease;
    cursor: pointer;
    position: relative;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.food-item:hover {
    transform: translateY(-3px);
    border-color: #3498db;
    box-shadow: 0 6px 20px rgba(52, 152, 219, 0.2);
}

.food-item.completed {
    background: linear-gradient(135deg, #d5f4e6 0%, #a8e6cf 100%);
    border-color: #27ae60;
}

.food-item.completed::after {
    content: '✓';
    position: absolute;
    top: 10px;
    right: 15px;
    color: #27ae60;
    font-size: 20px;
    font-weight: bold;
}

.food-name {
    font-weight: 700;
    color: #2c3e50;
    margin-bottom: 8px;
    font-size: 16px;
    line-height: 1.3;
}

.food-quantity {
    color: #3498db;
    font-size: 24px;
    font-weight: 700;
    margin-bottom: 5px;
}

.food-details {
    font-size: 13px;
    color: #7f8c8d;
    margin-top: 5px;
}

/* Stili per intolleranze se presenti */
.food-intolleranza {
    background: linear-gradient(135deg, #ffeaa7 0%, #fdcb6e 100%);
    border-color: #e17055;
}

.food-intolleranza .food-name::before {
    content: '🌱 ';
}
/* ===== CHECKBOX COMPLETAMENTO ===== */
.completion-checkbox {
    position: absolute;
    top: 15px;
    right: 15px;
    width: 25px;
    height: 25px;
    background: white;
    border: 3px solid #bdc3c7;
    border-radius: 6px;
    cursor: pointer;
    transition: all 0.3s ease;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: bold;
    color: white;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.completion-checkbox:hover {
    border-color: #27ae60;
    transform: scale(1.1);
}

.completion-checkbox.completed {
    background: #27ae60;
    border-color: #27ae60;
}

.completion-checkbox.completed::after {
    content: '✓';
    font-size: 16px;
}

.order-card.completed {
    background: linear-gradient(135deg, #d5f4e6 0%, #a8e6cf 100%);
    border-color: #27ae60;
}

.food-item .completion-checkbox {
    top: 10px;
    right: 10px;
    width: 20px;
    height: 20px;
    font-size: 12px;
}
    </style>
</head>
<body>
    <!-- HEADER -->
    <div class="header">
        <h1>🍳 Monitor Cucina</h1>
        <div class="company-subtitle">Ristorante Neri - Sistema v3.0</div>
        
        <div class="status">
            <div class="status-item" id="connection-status">
                <div class="status-dot"></div>
                <span id="connection-text">Connessione...</span>
            </div>
            <div class="status-item">
                <span>📋 Ordini: <strong id="order-count">0</strong></span>
            </div>
            <div class="status-item">
                <span>📅 Date: <strong id="date-count">0</strong></span>
            </div>
            <div class="status-item">
                <span>🕒 Aggiornato: <strong id="last-update">--:--</strong></span>
            </div>
        </div>
    </div>

    <div class="container">
        <!-- NAVIGATION -->
        <div class="navigation">
            <button class="back-button" id="back-button">← Torna alla Lista</button>
            <div class="breadcrumb" id="breadcrumb">Lista Ordini</div>
        </div>

        <!-- FILTERS -->
        <div class="filters" id="filters">
            <div class="filters-row">
                <div class="filter-group">
                    <label class="filter-label">Data</label>
                    <input type="date" class="filter-input" id="filter-date">
                </div>
                <div class="filter-group">
                    <label class="filter-label">Stato</label>
                    <select class="filter-input" id="filter-stato">
                        <option value="">Tutti</option>
                        <option value="attivo">Attivo</option>
                        <option value="completato">Completato</option>
                    </select>
                </div>
                <button class="filter-button" id="filter-today">📅 Solo Oggi</button>
                <button class="filter-button" id="filter-da-fare">📝 Da Fare</button>
                <button class="filter-button" id="show-totals">📊 Totali</button>
                <button class="filter-button" id="clear-filters">🔄 Reset</button>
            </div>
            
            <!-- SELECTION MODE BAR -->
            <div class="selection-bar" id="selection-bar" style="display: none;">
                <div class="selection-info">
                    <span id="selection-count">0 ordini selezionati</span>
                </div>
                <div class="selection-actions">
                    <button class="action-button" id="calculate-totals">🧮 Calcola Totali</button>
                    <button class="action-button secondary" id="cancel-selection">❌ Annulla</button>
                </div>
            </div>
        </div>

        <!-- LIST VIEW -->
        <div id="list-view">
            <div class="orders-grid" id="orders-container">
                <!-- Gli ordini verranno inseriti qui -->
            </div>
        </div>

        <!-- DETAIL VIEW -->
        <div class="detail-view" id="detail-view">
            <div class="detail-header">
                <div class="detail-title" id="detail-title">Dettagli Ordine</div>
                <div class="detail-info" id="detail-info">
                    <!-- Info ordine verranno inserite qui -->
                </div>
            </div>
            
            <div class="food-section">
                <div class="section-title">🍽️ Prodotti Ordinati</div>
                <div class="food-grid" id="food-container">
                    <!-- Prodotti verranno inseriti qui -->
                </div>
            </div>
        </div>
        <!-- TOTALS VIEW -->
        <div class="totals-view" id="totals-view" style="display: none;">
            <div class="totals-header">
                <h2 class="totals-title">📊 Totali Prodotti</h2>
                <div class="totals-info" id="totals-info">
                    <!-- Info totali verranno inserite qui -->
                </div>
            </div>
            
            <div class="totals-content">
                <div class="totals-grid" id="totals-container">
                    <!-- Totali prodotti verranno inseriti qui -->
                </div>
            </div>
        </div>

        <!-- LOADING -->
        <div class="loading" id="loading" style="display: none;">
            <div class="spinner"></div>
            <div>Caricamento ordini...</div>
        <!-- LOADING -->
        <div class="loading" id="loading" style="display: none;">
            <div class="spinner"></div>
            <div>Caricamento ordini...</div>
        </div>
    </div>

    <!-- FOOTER -->
    <div class="footer">
        <p>🔥 Connesso a <span id="server-info">Firebase Cloud</span> • Aggiornamento automatico ogni 30 secondi</p>
    </div>

    <script>
        // ===== CONFIGURAZIONE =====
        const CONFIG = {
            // Firebase
            firebase: {
                apiKey: "AIzaSyBz5RJrsSTBXCr6pU4uOikL0iac9-95Og0",
                authDomain: "neri-17bbd.firebaseapp.com",
                projectId: "neri-17bbd",
                storageBucket: "neri-17bbd.firebasestorage.app",
                messagingSenderId: "502407130695",
                appId: "1:502407130695:web:83a5406310e306ee274b62"
            },
            
            // Impostazioni
            useFirebase: true,
            serverUrl: 'http://localhost:5001',
            updateInterval: 30000, // 30 secondi
            
            // Collections Firebase
            collections: {
                ordiniSingoli: 'ordini_singoli',
                completedItems: 'completed_items',
                monitorState: 'monitor_state'
            }
        };

        // ===== CLASSE PRINCIPALE MONITOR =====
        class MonitorCucina {
            constructor() {
                this.ordini = [];
                this.filteredOrders = [];
                this.currentView = 'list'; // 'list', 'detail', 'totals'
                this.currentOrderId = null;
                this.isLoading = false;
                this.isConnected = false;
                this.completedProducts = new Set();
                this.initProductCompletion();
                
                // Stato filtri
                this.filters = {
                    data: '',
                    stato: '',
                    oggi: false,
                    daFare: false
                };
                
                // Stato selezione
                this.selectionMode = false;
                this.selectedOrders = new Set();
                
                // Firebase
                this.db = null;
                this.firebaseListener = null;

                // Stato completamento
                this.completedOrders = new Set();
                this.completedFoodItems = new Map();

                // Inizializza stato completamento
                this.initCompletionState();
                
                // Inizializza
                this.init();
            }

            async init() {
                console.log('🚀 Inizializzazione Monitor Cucina v3.0');
                
                try {
                    // Inizializza Firebase
                    await this.initFirebase();
                    
                    // Setup eventi DOM
                    this.setupEventListeners();
                    
                    // Carica ordini iniziali
                    await this.loadOrders();
                    
                    // Avvia aggiornamenti automatici
                    this.startAutoUpdate();
                    
                    console.log('✅ Monitor inizializzato con successo');
                    
                } catch (error) {
                    console.error('❌ Errore inizializzazione:', error);
                    this.showError('Errore durante l\'inizializzazione');
                }
            }

            async initFirebase() {
                if (!CONFIG.useFirebase) {
                    console.log('📡 Modalità server locale');
                    return;
                }

                try {
                    // Inizializza Firebase
                    firebase.initializeApp(CONFIG.firebase);
                    this.db = firebase.firestore();
                    
                    console.log('🔥 Firebase inizializzato');
                    
                    // Test connessione
                    await this.testFirebaseConnection();
                    
                } catch (error) {
                    console.error('❌ Errore Firebase:', error);
                    throw error;
                }
            }

            async testFirebaseConnection() {
                try {
                    // Test lettura collezione
                    const snapshot = await this.db.collection(CONFIG.collections.ordiniSingoli)
                        .limit(1)
                        .get();
                    
                    console.log('✅ Connessione Firebase OK');
                    this.setConnectionStatus(true, 'Firebase Online');
                    
                } catch (error) {
                    console.error('❌ Test Firebase fallito:', error);
                    this.setConnectionStatus(false, 'Errore Firebase');
                    throw error;
                }
            }

            setupEventListeners() {
                // Bottoni navigazione
                document.getElementById('back-button').addEventListener('click', () => {
                    this.showListView();
                });

                // Filtri
                document.getElementById('filter-date').addEventListener('change', (e) => {
                    this.filters.data = e.target.value;
                    this.applyFilters();
                });

                document.getElementById('filter-stato').addEventListener('change', (e) => {
                    this.filters.stato = e.target.value;
                    this.applyFilters();
                });

                document.getElementById('filter-today').addEventListener('click', () => {
                    this.toggleTodayFilter();
                });

                document.getElementById('filter-da-fare').addEventListener('click', () => {
                    this.toggleDaFareFilter();
                });

                document.getElementById('show-totals').addEventListener('click', () => {
                    this.showTotalsView();
                });

                document.getElementById('clear-filters').addEventListener('click', () => {
                    this.clearFilters();
                });

                // Selezione
                document.getElementById('calculate-totals').addEventListener('click', () => {
                    if (this.selectedOrders.size === 0) {
                        alert('Seleziona almeno un ordine per vedere i totali');
                        return;
                    }
                    this.displayTotalsView();
                });

                document.getElementById('cancel-selection').addEventListener('click', () => {
                    this.exitSelectionMode();
                });

                console.log('🎯 Event listeners configurati');
            }


            // ===== GESTIONE CONNESSIONE =====
            setConnectionStatus(connected, message) {
                this.isConnected = connected;
                const statusElement = document.getElementById('connection-status');
                const textElement = document.getElementById('connection-text');
                
                if (connected) {
                    statusElement.classList.remove('error');
                    textElement.textContent = message;
                } else {
                    statusElement.classList.add('error');
                    textElement.textContent = message;
                }
            }

            showError(message) {
                console.error('💥 Errore:', message);
                // Qui potresti aggiungere un toast notification
            }

            // ===== UTILITÁ DATE =====
            getTodayString() {
                const today = new Date();
                const day = String(today.getDate()).padStart(2, '0');
                const month = String(today.getMonth() + 1).padStart(2, '0');
                const year = today.getFullYear();
                return `${day}/${month}/${year}`;
            }

            formatTime() {
                const now = new Date();
                return now.toLocaleTimeString('it-IT', { 
                    hour: '2-digit', 
                    minute: '2-digit' 
                });
            }

            updateLastUpdate() {
                document.getElementById('last-update').textContent = this.formatTime();
            }
            // ===== CARICAMENTO ORDINI =====
            async loadOrders() {
                if (this.isLoading) {
                    console.log('⏳ Caricamento già in corso...');
                    return;
                }

                this.isLoading = true;
                this.showLoading(true);

                try {
                    const orderManager = new OrderManager(this);
                    
                    // Carica da Firebase o Server
                    if (CONFIG.useFirebase) {
                        this.ordini = await orderManager.loadFromFirebase();
                        this.setConnectionStatus(true, 'Firebase Online');
                    } else {
                        this.ordini = await orderManager.loadFromServer();
                        this.setConnectionStatus(true, 'Server Online');
                    }

                    // Aggiorna statistiche
                    this.updateStats();
                    
                    // Applica filtri e mostra ordini
                    this.applyFilters();
                    
                    console.log(`✅ Caricati ${this.ordini.length} ordini`);
                    
                } catch (error) {
                    console.error('❌ Errore caricamento ordini:', error);
                    this.setConnectionStatus(false, 'Errore Connessione');
                    this.showError('Impossibile caricare gli ordini');
                } finally {
                    this.isLoading = false;
                    this.showLoading(false);
                    this.updateLastUpdate();
                }
            }
                        // ===== GESTIONE COMPLETAMENTO PRODOTTI TOTALI =====
            initProductCompletion() {
                const saved = localStorage.getItem('completedProducts');
                this.completedProducts = saved ? new Set(JSON.parse(saved)) : new Set();
            }

            saveProductCompletion() {
                localStorage.setItem('completedProducts', JSON.stringify([...this.completedProducts]));
            }

            isProductCompleted(productName) {
                return this.completedProducts.has(productName);
            }

            toggleProductCompletion(productName) {
                if (this.completedProducts.has(productName)) {
                    this.completedProducts.delete(productName);
                } else {
                    this.completedProducts.add(productName);
                }
                
                this.saveProductCompletion();
                
                // Controlla se tutti i prodotti sono completati
                this.checkAllProductsCompleted();
                
                // Aggiorna vista totali
                this.renderTotalsView();
                
                console.log(`Prodotto ${productName} ${this.completedProducts.has(productName) ? 'completato' : 'riattivato'}`);
            }

            checkAllProductsCompleted() {
                // Ottieni tutti i prodotti dalla vista totali corrente
                const selectedOrders = this.ordini.filter(order => this.selectedOrders.has(order.id));
                const totals = this.renderer.calculateTotals(selectedOrders);
                
                // Controlla se tutti i prodotti sono completati
                const allCompleted = totals.products.every(product => 
                    this.completedProducts.has(product.nome)
                );
                
                if (allCompleted && totals.products.length > 0) {
                    // Completa tutti gli ordini selezionati
                    this.selectedOrders.forEach(orderId => {
                        this.completedOrders.add(orderId);
                    });
                    
                    this.saveCompletionState();
                    
                    alert(`Tutti i prodotti completati! Gli ordini selezionati (${this.selectedOrders.size}) sono stati marcati come completati.`);
                }
            }

            // ===== GESTIONE FILTRI =====
            applyFilters() {
                let filtered = [...this.ordini];

                // Filtro data
                if (this.filters.data) {
                    filtered = filtered.filter(order => 
                        order.data_evento === this.convertDateFormat(this.filters.data)
                    );
                }

                // Filtro oggi
                if (this.filters.oggi) {
                    const oggi = this.getTodayString();
                    filtered = filtered.filter(order => order.data_evento === oggi);
                }

                // Filtro stato
                if (this.filters.stato) {
                    filtered = filtered.filter(order => order.stato === this.filters.stato);
                }

                // Filtro da fare
                if (this.filters.daFare) {
                    filtered = filtered.filter(order => this.hasIncompletedItems(order));
                }

                this.filteredOrders = filtered;
                
                // Aggiorna vista
                if (this.currentView === 'list') {
                    this.renderOrdersList();
                } else if (this.currentView === 'totals') {
                    this.renderTotalsView();
                }

                console.log(`🔍 Filtri applicati: ${filtered.length}/${this.ordini.length} ordini`);
            }

            toggleTodayFilter() {
                this.filters.oggi = !this.filters.oggi;
                const btn = document.getElementById('filter-today');
                
                if (this.filters.oggi) {
                    btn.classList.add('active');
                    btn.textContent = '📅 Solo Oggi ✓';
                    // Reset filtro data (incompatibile)
                    this.filters.data = '';
                    document.getElementById('filter-date').value = '';
                } else {
                    btn.classList.remove('active');
                    btn.textContent = '📅 Solo Oggi';
                }
                
                this.applyFilters();
            }

            toggleDaFareFilter() {
                this.filters.daFare = !this.filters.daFare;
                const btn = document.getElementById('filter-da-fare');
                
                if (this.filters.daFare) {
                    btn.classList.add('active', 'da-fare');
                    btn.textContent = '📝 Da Fare ✓';
                } else {
                    btn.classList.remove('active', 'da-fare');
                    btn.textContent = '📝 Da Fare';
                }
                
                this.applyFilters();
            }

            clearFilters() {
                this.filters = {
                    data: '',
                    stato: '',
                    oggi: false,
                    daFare: false
                };

                // Reset UI
                document.getElementById('filter-date').value = '';
                document.getElementById('filter-stato').value = '';
                document.getElementById('filter-today').classList.remove('active');
                document.getElementById('filter-today').textContent = '📅 Solo Oggi';
                document.getElementById('filter-da-fare').classList.remove('active', 'da-fare');
                document.getElementById('filter-da-fare').textContent = '📝 Da Fare';

                this.applyFilters();
            }

            // ===== GESTIONE VISTE =====
            showListView() {
                this.currentView = 'list';
                this.currentOrderId = null;
                
                // Esci da modalità selezione se era attiva per i totali
                if (this.selectionMode) {
                    this.exitSelectionMode();
                }
                
                // Mostra/nascondi elementi
                document.getElementById('list-view').style.display = 'block';
                document.getElementById('detail-view').style.display = 'none';
                document.getElementById('totals-view').style.display = 'none';
                document.getElementById('back-button').style.display = 'none';
                document.getElementById('filters').style.display = 'block';
                
                // Aggiorna breadcrumb
                document.getElementById('breadcrumb').textContent = 'Lista Ordini';
                
                this.renderOrdersList();
            }

            showDetailView(orderId) {
                this.currentView = 'detail';
                this.currentOrderId = orderId;
                
                // Mostra/nascondi elementi
                document.getElementById('list-view').style.display = 'none';
                document.getElementById('detail-view').style.display = 'block';
                document.getElementById('totals-view').style.display = 'none';
                document.getElementById('back-button').style.display = 'block';
                document.getElementById('filters').style.display = 'none';
                
                // Trova ordine
                const order = this.ordini.find(o => o.id === orderId);
                if (order) {
                    document.getElementById('breadcrumb').textContent = 
                        `Dettagli: ${order.nome_cliente}`;
                    this.renderOrderDetail(order);
                }
            }

            showTotalsView() {
                    // Entra in modalità selezione per scegliere ordini
                    if (!this.selectionMode) {
                        this.enterSelectionMode();
                        
                        // Mostra messaggio informativo
                        alert('Modalità selezione attivata!\n\nSeleziona gli ordini di cui vuoi vedere i totali cliccando sulle card,\npoi premi "Mostra Totali" per visualizzare il riepilogo.');
                        
                        // Cambia il testo del pulsante nella barra di selezione
                        document.getElementById('calculate-totals').textContent = '📊 Mostra Totali';
                        
                        return;
                    }
                    
                    // Se siamo già in modalità selezione e hanno selezionato ordini
                    if (this.selectedOrders.size > 0) {
                        this.displayTotalsView();
                    } else {
                        alert('Seleziona almeno un ordine per vedere i totali');
                    }
            }
            displayTotalsView() {
                this.currentView = 'totals';
                
                // Mostra/nascondi elementi
                document.getElementById('list-view').style.display = 'none';
                document.getElementById('detail-view').style.display = 'none';
                document.getElementById('totals-view').style.display = 'block';
                document.getElementById('back-button').style.display = 'block';
                document.getElementById('filters').style.display = 'none';
                
                // Nascondi barra selezione
                document.getElementById('selection-bar').style.display = 'none';
                
                // Aggiorna breadcrumb
                document.getElementById('breadcrumb').textContent = `Totali (${this.selectedOrders.size} ordini)`;
                
                this.renderTotalsView();
            }

            // ===== GESTIONE SELEZIONE =====
            enterSelectionMode() {
                this.selectionMode = true;
                this.selectedOrders.clear();
                
                // Mostra barra selezione
                document.getElementById('selection-bar').style.display = 'flex';
                
                // Aggiorna vista ordini
                this.renderOrdersList();
                
                this.updateSelectionCount();
            }

            exitSelectionMode() {
                this.selectionMode = false;
                this.selectedOrders.clear();
                
                // Nascondi barra selezione
                document.getElementById('selection-bar').style.display = 'none';
                
                // Aggiorna vista ordini
                this.renderOrdersList();
            }

            toggleOrderSelection(orderId) {
                if (this.selectedOrders.has(orderId)) {
                    this.selectedOrders.delete(orderId);
                } else {
                    this.selectedOrders.add(orderId);
                }
                
                this.updateSelectionCount();
                this.renderOrdersList();
            }

            updateSelectionCount() {
                const count = this.selectedOrders.size;
                document.getElementById('selection-count').textContent = 
                    `${count} ordini selezionati`;
            }

            // ===== UTILITIES =====
            convertDateFormat(dateString) {
                // Converte da YYYY-MM-DD a DD/MM/YYYY
                if (!dateString) return '';
                const [year, month, day] = dateString.split('-');
                return `${day}/${month}/${year}`;
            }

            hasIncompletedItems(order) {
                // Verifica se l'ordine ha prodotti non completati
                if (!order.cibo) return false;
                
                // Per ora restituisce true se ha prodotti
                // TODO: implementare logica completamento
                return Object.keys(order.cibo).length > 0;
            }

            updateStats() {
                const totalOrders = this.ordini.length;
                const uniqueDates = new Set(
                    this.ordini
                        .map(o => o.data_evento)
                        .filter(d => d)
                ).size;

                document.getElementById('order-count').textContent = totalOrders;
                document.getElementById('date-count').textContent = uniqueDates;
            }
            // ===== GESTIONE COMPLETAMENTO =====
            initCompletionState() {
                // Carica stato completamento da localStorage
                const savedCompleted = localStorage.getItem('completedOrders');
                const savedFoodItems = localStorage.getItem('completedFoodItems');
                
                this.completedOrders = savedCompleted ? new Set(JSON.parse(savedCompleted)) : new Set();
                this.completedFoodItems = savedFoodItems ? new Map(JSON.parse(savedFoodItems)) : new Map();
            }

            saveCompletionState() {
                // Salva stato in localStorage
                localStorage.setItem('completedOrders', JSON.stringify([...this.completedOrders]));
                localStorage.setItem('completedFoodItems', JSON.stringify([...this.completedFoodItems]));
            }

            isOrderCompleted(orderId) {
                return this.completedOrders.has(orderId);
            }

            isFoodItemCompleted(orderId, foodName) {
                const orderItems = this.completedFoodItems.get(orderId);
                return orderItems ? orderItems.has(foodName) : false;
            }

            toggleOrderCompletion(orderId) {
                if (this.completedOrders.has(orderId)) {
                    this.completedOrders.delete(orderId);
                    // Rimuovi anche tutti i prodotti dell'ordine
                    this.completedFoodItems.delete(orderId);
                } else {
                    this.completedOrders.add(orderId);
                    // Completa automaticamente tutti i prodotti
                    const order = this.ordini.find(o => o.id === orderId);
                    if (order && order.cibo) {
                        const foodItems = new Set(Object.keys(order.cibo));
                        this.completedFoodItems.set(orderId, foodItems);
                    }
                }
                
                this.saveCompletionState();
                this.renderOrdersList();
                console.log(`🔄 Ordine ${orderId} ${this.completedOrders.has(orderId) ? 'completato' : 'riattivato'}`);
            }

            toggleFoodItemCompletion(orderId, foodName) {
                if (!this.completedFoodItems.has(orderId)) {
                    this.completedFoodItems.set(orderId, new Set());
                }
                
                const orderItems = this.completedFoodItems.get(orderId);
                if (orderItems.has(foodName)) {
                    orderItems.delete(foodName);
                } else {
                    orderItems.add(foodName);
                }
                
                // Controlla se tutti i prodotti sono completati
                const order = this.ordini.find(o => o.id === orderId);
                if (order && order.cibo) {
                    const totalItems = Object.keys(order.cibo).length;
                    const completedItems = orderItems.size;
                    
                    if (completedItems === totalItems && totalItems > 0) {
                        // Tutti i prodotti completati -> completa ordine
                        this.completedOrders.add(orderId);
                    } else {
                        // Non tutti completati -> rimuovi completamento ordine
                        this.completedOrders.delete(orderId);
                    }
                }
                
                this.saveCompletionState();
                
                // Aggiorna vista
                if (this.currentView === 'detail') {
                    const order = this.ordini.find(o => o.id === orderId);
                    this.renderOrderDetail(order);
                } else {
                    this.renderOrdersList();
                }
                
                console.log(`🍽️ Prodotto ${foodName} in ordine ${orderId} ${orderItems.has(foodName) ? 'completato' : 'riattivato'}`);
            }

            showLoading(show) {
                document.getElementById('loading').style.display = show ? 'block' : 'none';
            }

            startAutoUpdate() {
                setInterval(() => {
                    if (!this.isLoading) {
                        this.loadOrders();
                    }
                }, CONFIG.updateInterval);
                
                console.log(`🔄 Auto-update attivato ogni ${CONFIG.updateInterval/1000}s`);
            }
        }

        // ===== CLASSE GESTIONE ORDINI =====
        class OrderManager {
            constructor(monitor) {
                this.monitor = monitor;
            }

            async loadFromFirebase() {
                if (!this.monitor.db) return [];

                try {
                    console.log('📥 Caricamento ordini da Firebase...');
                    
                    const snapshot = await this.monitor.db
                        .collection(CONFIG.collections.ordiniSingoli)
                        .get();
                    
                    const ordini = [];
                    const ordersMap = new Map();
                    
                    snapshot.forEach(doc => {
                        const data = doc.data(); // ✅ CORREZIONE: era doc.to_dict()
                        const orderId = doc.id;
                        
                        // Evita duplicati
                        if (!ordersMap.has(orderId)) {
                            const order = this.processOrderData(orderId, data);
                            if (order) {
                                ordersMap.set(orderId, order);
                                ordini.push(order);
                            }
                        }
                    });
                    
                    console.log(`✅ Caricati ${ordini.length} ordini da Firebase`);
                    return ordini;
                    
                } catch (error) {
                    console.error('❌ Errore caricamento Firebase:', error);
                    throw error;
                }
            }

            async loadFromServer() {
                try {
                    console.log('📥 Caricamento ordini da server locale...');
                    
                    const response = await fetch(`${CONFIG.serverUrl}/api/ordini`);
                    if (!response.ok) {
                        throw new Error(`HTTP ${response.status}`);
                    }
                    
                    const data = await response.json();
                    const ordini = data.ordini || [];
                    
                    console.log(`✅ Caricati ${ordini.length} ordini da server`);
                    return ordini;
                    
                } catch (error) {
                    console.error('❌ Errore caricamento server:', error);
                    throw error;
                }
            }

            processOrderData(orderId, data) {
                try {
                    // Valida dati base
                    if (!data.info || !data.info.nome_cliente) {
                        console.warn(`⚠️ Ordine ${orderId}: dati info mancanti`);
                        return null;
                    }

                    // Conta prodotti
                    const prodottiCount = data.cibo ? 
                        Object.keys(data.cibo).length : 0;

                    return {
                        id: orderId,
                        nome_cliente: data.info.nome_cliente,
                        data_evento: data.info.data_evento || null,
                        filename: data.info.filename || '',
                        prodotti_count: prodottiCount,
                        stato: data.stato || 'attivo',
                        ultimo_aggiornamento: data.ultimo_aggiornamento || new Date().toISOString(),
                        cibo: data.cibo || {},
                        info: data.info
                    };
                } catch (error) {
                    console.error(`❌ Errore elaborazione ordine ${orderId}:`, error);
                    return null;
                }
            }
        }

        // ===== CLASSE RENDERING =====
        class UIRenderer {
            constructor(monitor) {
                this.monitor = monitor;
            }

            renderOrdersList() {
                const container = document.getElementById('orders-container');
                
                if (this.monitor.filteredOrders.length === 0) {
                    container.innerHTML = `
                        <div style="grid-column: 1/-1; text-align: center; padding: 50px; color: white;">
                            <h3>📭 Nessun ordine trovato</h3>
                            <p>Prova a modificare i filtri o ricarica la pagina</p>
                        </div>
                    `;
                    return;
                }

                const ordersHtml = this.monitor.filteredOrders.map(order => {
                    const isSelected = this.monitor.selectedOrders.has(order.id);
                    const selectionClass = this.monitor.selectionMode ? 'selection-mode' : '';
                    const selectedClass = isSelected ? 'selected' : '';
                    const isCompleted = this.monitor.isOrderCompleted(order.id);
                    const completedClass = isCompleted ? 'completed' : '';
                    
                    return `
                        <div class="order-card ${selectionClass} ${selectedClass} ${completedClass}" 
                            data-order-id="${order.id}" style="position: relative;">
                            
                            <!-- Checkbox completamento ordine (nascosta se in modalità selezione) -->
                            ${this.monitor.selectionMode ? '' : `
                                <div class="completion-checkbox ${isCompleted ? 'completed' : ''}" 
                                    onclick="event.stopPropagation(); window.monitor.toggleOrderCompletion('${order.id}')">
                                </div>
                            `}
                            
                            <div class="order-header">
                                <div class="order-title">${order.nome_cliente}</div>
                                ${this.renderOrderBadges(order)}
                            </div>
                            
                            <div class="order-info">
                                <div class="info-row">
                                    <span class="info-label">📅 Data:</span>
                                    <span class="info-value">${order.data_evento || 'Non specificata'}</span>
                                </div>
                                <div class="info-row">
                                    <span class="info-label">🍽️ Prodotti:</span>
                                    <span class="info-value">${order.prodotti_count}</span>
                                </div>
                                <div class="info-row">
                                    <span class="info-label">📄 File:</span>
                                    <span class="info-value">${order.filename}</span>
                                </div>
                                <div class="info-row">
                                    <span class="info-label">🔄 Stato:</span>
                                    <span class="info-value">${this.formatStato(order.stato)}</span>
                                </div>
                            </div>
                            
                            ${order.prodotti_count > 0 ? `
                                <div class="order-products">
                                    <div class="products-summary">
                                        Clicca per vedere i ${order.prodotti_count} prodotti
                                    </div>
                                </div>
                            ` : ''}
                        </div>
                    `;
                }).join('');

                container.innerHTML = ordersHtml;
                
                // Event listeners per click sulle card
                container.querySelectorAll('.order-card').forEach(card => {
                    card.addEventListener('click', (e) => {
                        // Ignora se clicchi sulla checkbox
                        if (e.target.classList.contains('completion-checkbox')) return;
                        
                        const orderId = card.dataset.orderId;
                        if (this.monitor.selectionMode) {
                            this.monitor.toggleOrderSelection(orderId);
                        } else {
                            this.monitor.showDetailView(orderId);
                        }
                    });
                });
            }

            renderOrderBadges(order) {
                const badges = [];
                
                // Badge nuovo (esempio - puoi implementare la logica)
                if (this.isNewOrder(order)) {
                    badges.push('<span class="order-badge badge-new">NUOVO</span>');
                }
                
                // Badge modificato (esempio - puoi implementare la logica)
                if (this.isModifiedOrder(order)) {
                    badges.push('<span class="order-badge badge-modified">MODIFICATO</span>');
                }
                
                return badges.join('');
            }

            renderOrderDetail(order) {
                // Aggiorna titolo
                document.getElementById('detail-title').textContent = 
                    `Ordine: ${order.nome_cliente}`;

                // Aggiorna info
                const infoContainer = document.getElementById('detail-info');
                infoContainer.innerHTML = `
                    <div class="info-item">
                        <div class="info-label">📅 Data Evento</div>
                        <div class="info-value">${order.data_evento || 'Non specificata'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">👥 Cliente</div>
                        <div class="info-value">${order.nome_cliente}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">📄 File</div>
                        <div class="info-value">${order.filename}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">🍽️ Prodotti</div>
                        <div class="info-value">${order.prodotti_count}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">🔄 Stato</div>
                        <div class="info-value">${this.formatStato(order.stato)}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">🕒 Aggiornato</div>
                        <div class="info-value">${this.formatDateTime(order.ultimo_aggiornamento)}</div>
                    </div>
                `;

                // Aggiorna prodotti
                this.renderFoodItems(order);
            }

            renderFoodItems(order) {
                const container = document.getElementById('food-container');
                
                if (!order.cibo || Object.keys(order.cibo).length === 0) {
                    container.innerHTML = `
                        <div style="grid-column: 1/-1; text-align: center; padding: 30px; color: #7f8c8d;">
                            <h4>📭 Nessun prodotto trovato</h4>
                            <p>Questo ordine non contiene prodotti alimentari</p>
                        </div>
                    `;
                    return;
                }

                // Separa prodotti normali e intolleranze
                const prodottiNormali = {};
                const intolleranze = {};
                
                Object.entries(order.cibo).forEach(([nome, quantita]) => {
                    if (nome.toLowerCase().includes('intolleranza') || 
                        nome.toLowerCase().includes('allergia') ||
                        nome.toLowerCase().includes('celiaco')) {
                        intolleranze[nome] = quantita;
                    } else {
                        prodottiNormali[nome] = quantita;
                    }
                });

                let foodHtml = '';

                // Renderizza prodotti normali
                // Renderizza prodotti normali
                if (Object.keys(prodottiNormali).length > 0) {
                    foodHtml += Object.entries(prodottiNormali).map(([nome, quantita]) => {
                        const isCompleted = this.monitor.isFoodItemCompleted(order.id, nome);
                        const completedClass = isCompleted ? 'completed' : '';
                        
                        // Gestisci quantità che può essere oggetto
                        let quantitaDisplay = quantita;
                        if (typeof quantita === 'object' && quantita !== null) {
                            quantitaDisplay = quantita.value || quantita.quantita || quantita.qty || JSON.stringify(quantita);
                        }
                        
                        return `
                            <div class="food-item ${completedClass}" style="position: relative;">
                                <!-- Checkbox completamento prodotto -->
                                <div class="completion-checkbox ${isCompleted ? 'completed' : ''}" 
                                    onclick="window.monitor.toggleFoodItemCompletion('${order.id}', '${nome.replace(/'/g, "\\'")}')">
                                </div>
                                
                                <div class="food-name">${nome}</div>
                                <div class="food-quantity">${quantitaDisplay}</div>
                                <div class="food-details">
                                    ${isCompleted ? 'Completato ✓' : 'Da completare'}
                                </div>
                            </div>
                        `;
                    }).join('');
                }

                // Renderizza intolleranze se presenti
                if (Object.keys(intolleranze).length > 0) {
                    foodHtml += `
                        <div style="grid-column: 1/-1;">
                            <div style="margin: 20px 0; text-align: center;">
                                <div style="background: rgba(230, 126, 34, 0.1); padding: 10px; border-radius: 8px; border-left: 4px solid #e67e22;">
                                    <h4 style="color: #e67e22; margin: 0;">🌱 INTOLLERANZE E ALLERGIE</h4>
                                </div>
                            </div>
                        </div>
                    `;
                    
                    foodHtml += Object.entries(intolleranze).map(([nome, quantita]) => {
                        // Gestisci quantità per intolleranze
                        let quantitaDisplay = quantita;
                        if (typeof quantita === 'object' && quantita !== null) {
                            quantitaDisplay = quantita.value || quantita.quantita || quantita.qty || JSON.stringify(quantita);
                        }

                        return `
                            <div class="food-item food-intolleranza">
                                <div class="food-name">${nome}</div>
                                <div class="food-quantity">${quantitaDisplay}</div>
                                <div class="food-details">Attenzione speciale richiesta</div>
                            </div>
                        `;
                    }).join('');
                }

                container.innerHTML = foodHtml;
            }
            renderTotalsView() {
                // Calcola totali SOLO dagli ordini selezionati
                const selectedOrders = this.monitor.ordini.filter(order => 
                    this.monitor.selectedOrders.has(order.id)
                );
                
                const totals = this.calculateTotals(selectedOrders);
                
                // Aggiorna info header
                const infoContainer = document.getElementById('totals-info');
                infoContainer.innerHTML = `
                    <div class="totals-stat">
                        <div class="stat-number">${totals.ordersCount}</div>
                        <div class="stat-label">Ordini Selezionati</div>
                    </div>
                    <div class="totals-stat">
                        <div class="stat-number">${totals.productsCount}</div>
                        <div class="stat-label">Prodotti Diversi</div>
                    </div>
                    <div class="totals-stat">
                        <div class="stat-number">${totals.totalQuantity}</div>
                        <div class="stat-label">Quantità Totale</div>
                    </div>
                `;

                // Mostra anche lista ordini selezionati
                const ordersList = selectedOrders.map(order => 
                    `• ${order.nome_cliente} (${order.data_evento || 'N/A'})`
                ).join('<br>');
                
                infoContainer.innerHTML += `
                    <div class="totals-stat" style="min-width: 300px;">
                        <div class="stat-label">Ordini inclusi:</div>
                        <div style="font-size: 12px; margin-top: 10px; text-align: left;">
                            ${ordersList}
                        </div>
                    </div>
                `;

                // Aggiorna griglia totali
                this.renderTotalsGrid(totals.products);
            }

            renderTotalsGrid(products) {
                const container = document.getElementById('totals-container');
                
                if (products.length === 0) {
                    container.innerHTML = `
                        <div style="grid-column: 1/-1; text-align: center; padding: 50px; color: #7f8c8d;">
                            <h3>📭 Nessun prodotto trovato</h3>
                            <p>Nessun ordine contiene prodotti o i filtri sono troppo restrittivi</p>
                        </div>
                    `;
                    return;
                }

                const totalsHtml = products.map(product => {
                    const isCompleted = this.monitor.isProductCompleted && this.monitor.isProductCompleted(product.nome);
                    return `
                        <div class="total-item ${isCompleted ? 'completed' : ''}" style="position: relative;">
                            <!-- Checkbox completamento prodotto -->
                            <div class="completion-checkbox ${isCompleted ? 'completed' : ''}" 
                                onclick="window.monitor.toggleProductCompletion('${product.nome.replace(/'/g, "\\'")}')">
                            </div>
                            
                            <div class="total-name">${product.nome}</div>
                            <div class="total-quantity">${product.quantita}</div>
                            <div class="total-orders">In ${product.ordini} ordini</div>
                        </div>
                    `;
                }).join('');

                container.innerHTML = totalsHtml;
            }

            calculateTotals(orders) {
                const totals = {};
                let totalQuantity = 0;

                orders.forEach(order => {
                    if (order.cibo) {
                        // La struttura è: order.cibo[key] = { nome: "...", quantita: "5 pz" }
                        Object.values(order.cibo).forEach(item => {
                            const name = item.nome;
                            const quantity = item.quantita;
                            
                            if (!totals[name]) {
                                totals[name] = { totalQuantity: 0, unit: '', ordini: 0 };
                            }
                            
                            // Stesso parsing della tua funzione
                            const match = quantity.match(/(\d+\.?\d*)\s*(.*)/);
                            if (match) {
                                const num = parseFloat(match[1]);
                                const unit = match[2];
                                totals[name].totalQuantity += num;
                                totals[name].unit = unit;
                                totals[name].ordini += 1;
                                totalQuantity += num;
                            }
                        });
                    }
                });

                // Converti in formato compatibile con il renderer
                const products = Object.entries(totals).map(([nome, data]) => ({
                    nome: nome,
                    quantita: `${data.totalQuantity} ${data.unit}`.trim(),
                    ordini: data.ordini
                }));

                // Ordina per quantità numerica decrescente
                products.sort((a, b) => {
                    const qtyA = parseFloat(a.quantita) || 0;
                    const qtyB = parseFloat(b.quantita) || 0;
                    return qtyB - qtyA;
                });

                return {
                    ordersCount: orders.length,
                    productsCount: products.length,
                    totalQuantity: Math.round(totalQuantity * 100) / 100, // Arrotonda a 2 decimali
                    products: products
                };
            }
            calculateSelectedTotals() {
            // Non mostra più alert, ma passa alla vista totali
            if (window.monitorInstance) {
                window.monitorInstance.displayTotalsView();
            }
        }

            // ===== UTILITIES =====
            formatStato(stato) {
                const stati = {
                    'attivo': '🟢 Attivo',
                    'completato': '✅ Completato',
                    'annullato': '❌ Annullato'
                };
                return stati[stato] || stato;
            }

            formatDateTime(isoString) {
                if (!isoString) return 'Non disponibile';
                try {
                    const date = new Date(isoString);
                    return date.toLocaleString('it-IT');
                } catch {
                    return 'Formato non valido';
                }
            }

            isNewOrder(order) {
                // Implementa logica per determinare se è un ordine nuovo
                // Per ora restituisce false
                return false;
            }

            isModifiedOrder(order) {
                // Implementa logica per determinare se è modificato
                // Per ora restituisce false
                return false;
            }
        }

        // ===== GESTORI EVENTI GLOBALI (DISPONIBILI SUBITO) =====
        window.monitor = {
            toggleProductCompletion(productName) {
    if (window.monitorInstance) {
        window.monitorInstance.toggleProductCompletion(productName);
    }
},
    handleOrderClick(orderId) {
        if (window.monitorInstance) {
            if (window.monitorInstance.selectionMode) {
                window.monitorInstance.toggleOrderSelection(orderId);
            } else {
                window.monitorInstance.showDetailView(orderId);
            }
        } else {
            console.warn('⚠️ Monitor non ancora inizializzato');
        }
    },

    toggleFoodItem(orderId, foodName) {
        if (window.monitorInstance) {
            console.log(`🍽️ Toggle prodotto: ${foodName} in ordine ${orderId}`);
            // TODO: implementare toggle completamento prodotto
        }
    },

    toggleOrderCompletion(orderId) {
        if (window.monitorInstance) {
            window.monitorInstance.toggleOrderCompletion(orderId);
        }
    },

    toggleFoodItemCompletion(orderId, foodName) {
        if (window.monitorInstance) {
            window.monitorInstance.toggleFoodItemCompletion(orderId, foodName);
        }
    },

    showListView() {
        if (window.monitorInstance) {
            window.monitorInstance.showListView();
        }
    },

    showTotalsView() {
        if (window.monitorInstance) {
            window.monitorInstance.showTotalsView();
        }
    },

    calculateSelectedTotals() {
        if (window.monitorInstance) {
            window.monitorInstance.showTotalsView();
        }
    }
};

// ===== INIZIALIZZAZIONE E GESTORI EVENTI =====
let monitor;
let renderer;

// Inizializzazione al caricamento pagina
document.addEventListener('DOMContentLoaded', async () => {
    console.log('🚀 Avvio Monitor Cucina v3.0');
    
    try {
        // Crea istanze
        monitor = new MonitorCucina();
        window.monitorInstance = monitor; // Rendi disponibile globalmente
        renderer = new UIRenderer(monitor);
        
        // Collega renderer al monitor
        monitor.renderer = renderer;
        
        console.log('✅ Monitor Cucina v3.0 inizializzato con successo');
        
    } catch (error) {
        console.error('💥 Errore fatale durante inizializzazione:', error);
        alert('Errore durante l\'inizializzazione del monitor. Controlla la console per dettagli.');
    }
});

// Aggiorna metodi del monitor per usare il renderer
MonitorCucina.prototype.renderOrdersList = function() {
    this.renderer.renderOrdersList();
};

MonitorCucina.prototype.renderOrderDetail = function(order) {
    this.renderer.renderOrderDetail(order);
};

MonitorCucina.prototype.renderTotalsView = function() {
    this.renderer.renderTotalsView();
};

        
    </script>
</body>
</html>