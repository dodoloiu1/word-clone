import './styles/main.css';
import { createWordApp } from './editor/wordApp.js';

const appRoot = document.getElementById('app');
if (appRoot) {
  window.wordApp = createWordApp(appRoot);
}
