import React from 'react';
import { Link } from 'react-router-dom';

const HeroNavbar = () => (
  <nav style={{ padding: '1rem', backgroundColor: '#222', color: '#fff' }}>
    <Link to="/" style={{ color: '#fff', marginRight: '1rem' }}>Home</Link>
    <Link to="/upload" style={{ color: '#fff', marginRight: '1rem' }}>Upload</Link>
    <Link to="/help" style={{ color: '#fff' }}>Help</Link>
  </nav>
);

export default HeroNavbar;
