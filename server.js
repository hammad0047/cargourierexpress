const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const session = require('express-session');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// Add middleware to parse JSON bodies
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure session middleware
app.use(session({
    secret: 'cargourierexpress_secret_key',
    resave: false,
    saveUninitialized: true,
    cookie: { 
        secure: false, // Set to true if using HTTPS
        maxAge: 24 * 60 * 60 * 1000 // 24 hours
    }
}));

// Serve static files from the current directory
app.use(express.static(__dirname));

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, __dirname);
    },
    filename: function (req, file, cb) {
        cb(null, 'performance_log.xlsx');
    }
});

const upload = multer({ storage: storage });

// Login endpoint
app.post('/login', (req, res) => {
    const { username, password } = req.body;
    
    // Check credentials
    if (username === 'admin' && password === 'admin123') {
        req.session.isAuthenticated = true;
        res.json({ success: true });
    } else {
        res.status(401).json({ success: false, message: 'Invalid credentials' });
    }
});

// Middleware to check authentication
const requireAuth = (req, res, next) => {
    if (req.session.isAuthenticated) {
        next();
    } else {
        res.status(401).json({ error: 'Unauthorized' });
    }
};

// Check authentication status
app.get('/check-auth', (req, res) => {
    res.json({ isAuthenticated: !!req.session.isAuthenticated });
});

// Upload endpoint
app.post('/upload', requireAuth, upload.single('file'), (req, res) => {
    res.json({ message: 'File uploaded successfully' });
});

// Performance data endpoint
app.get('/performance-data', requireAuth, (req, res) => {
    try {
        const workbook = XLSX.readFile(path.join(__dirname, 'performance_log.xlsx'));
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        res.json(data);
    } catch (error) {
        console.error('Error reading Excel file:', error);
        res.status(500).json({ error: 'Error reading Excel file' });
    }
});

// Logout endpoint
app.post('/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

// Route for the dashboard
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, () => {
    console.log(`Dashboard server running at http://localhost:${port}`);
}); 