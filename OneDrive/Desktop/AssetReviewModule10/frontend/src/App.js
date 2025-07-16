// src/App.js
import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import HeroNavbar from './components/HeroNavbar';
import HeroSection from './components/HeroSection';
import Upload from './components/Upload';
import Help from './components/Help';

function App() {
  return (
    <Router>
      <HeroNavbar />
      <Routes>
        <Route path="/" element={<HeroSection />} />
        <Route path="/upload" element={<Upload />} />
        <Route path="/help" element={<Help />} />
      </Routes>
    </Router>
  );
}

export default App;

