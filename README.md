# 🎮 Office Game - Hack&Roll 2026

A fully playable arcade game collection built entirely in **Microsoft Excel VBA**! Battle through classic Pong with dynamic barriers and cracking walls, then unlock a bonus Flappy Bird game through an epic egg-hatching cutscene.

## 🎯 Project Overview

This project was created for **Hack&Roll 2026** (24-hour hackathon) and showcases the unexpected power of Excel as a gaming platform. No external libraries, no game engines—just pure VBA magic!

### Games Included:
1. **Pong** - Classic paddle game with modern twists
2. **Flappy Bird** - Navigate through pipes with pixel-perfect collision
3. **Animated Cutscene** - Smooth transition animation between games

## ✨ Features

### Pong Game
- 🏓 **Classic Paddle Mechanics** - Move up/down to keep the ball in play
- 🧱 **Dynamic Barriers** - Obstacles spawn and degrade over time
  - 2-4 cell variable lengths
  - Horizontal & vertical orientations (80% vertical spawn rate)
  - Health-based color degradation
- 🧱 **Breakable Wall System** - Hit the wall 5 times to break through
  - Progressive crack visualization
  - 2-cell thick wall with individual cell damage
- 🎨 **Ball Trail Effect** - Smooth motion trail behind the ball
- 📈 **Difficulty Scaling** - Ball speed increases over time
- 🐉 **Pixel Art Dragon** - Hand-drawn boss behind the wall

### Flappy Bird Game
- 🐦 **Smooth Flight Mechanics** - Gravity + flap physics
- 🎨 **Custom Pipe Design** - Hand-drawn pipe pixel art (replicated from template)
- 📊 **Score Tracking** - Points for each pipe passed
- ⚡ **Adjustable Difficulty** - Configurable gap size, speed, and spawn rate
- 🎮 **One-Button Control** - Simple flap button gameplay

### Cutscene Animation
- 🥚 **Egg Hatching Sequence**:
  1. Ball flies to center stage
  2. Ball grows into egg (1x1 → 2x2 → 3x3 → oval)
  3. Cracks appear progressively
  4. Egg shakes and explodes
  5. Bird emerges and flies
- 💥 **Wall Explosion** - Wall pieces fly outward with physics (gravity + velocity)
- 🎬 **Smooth Frame Animation** - Cell-based flipbook animation

## 🎮 How to Play

### Setup
1. Open `PongGame.xlsm` in Microsoft Excel (macros enabled)
2. Go to the **Menu** sheet
3. Click **"Start Game"** to begin Pong

### Pong Controls
- **Up Button** (or click cell) - Move paddle up
- **Down Button** (or click cell) - Move paddle down
- **Objective**: Hit the wall 5 times to break through and win!

### Flappy Bird Controls
- **FLAP Button** - Make the bird jump
- **Objective**: Navigate through pipes and get the highest score!

### Game Flow
```
Menu → Pong → Wall Breaks → Explosion Animation → 
Egg Hatching → Bird Emerges → Flappy Bird → Game Over → Menu
```

## 🏗️ Technical Architecture

### File Structure
```
PongGame.xlsm
├── Sheets
│   ├── Menu (Game selection)
│   ├── Pong (Main game board)
│   └── FlappyBird (Bonus game board)
└── VBA Modules
    ├── Module1 (Pong game logic)
    ├── Module2 (Cutscene animations)
    └── Module3 (Flappy Bird logic)
```

### Key Technologies
- **VBA (Visual Basic for Applications)** - All game logic
- **Excel Cells as Pixels** - Visual rendering system
- **Timer-Based Game Loop** - `Application.OnTime` for smooth animation
- **RGB Color Manipulation** - Custom color palettes and effects
- **Cell Interior Colors** - Graphics rendering

### Core Systems

#### 1. Game Loop System
```vba
' Pong runs at 0.5 seconds per tick
' Flappy Bird runs at 0.15 seconds per tick
Application.OnTime GameTimer, "GameTick"
```

#### 2. Collision Detection
- **Pong**: Ball vs Paddle, Ball vs Barriers, Ball vs Walls
- **Flappy Bird**: Bird vs Pipes, Bird vs Ground, Bird vs Ceiling
- Pixel-perfect hitbox calculations

#### 3. Animation System
- **Frame-based animation** using DoEvents loops
- **Physics simulation** (gravity, velocity, friction)
- **Particle system** for wall explosion

## 🎨 Design Decisions

### Why Excel?
- **Accessibility** - Everyone has Excel, no installation needed
- **Challenge** - Pushing Excel beyond its intended use
- **Visual Grid** - Perfect for pixel art and retro games
- **Hackathon Novelty** - Unique approach for a 24-hour sprint

### Performance Optimizations
- `Application.ScreenUpdating = False` during rendering
- Efficient cell range clearing (batch operations)
- Minimal redraw operations (only changed cells)
- Timer-based game loop (non-blocking)


## 🐛 Known Issues & Limitations

### Performance
- ⚠️ Laggy on older computers (Excel isn't optimized for gaming!)
- ⚠️ Animation framerate depends on CPU speed
- ⚠️ Large number of active timers can cause slowdown

### Excel Quirks
- ⚠️ `Application.OnTime` can sometimes queue multiple callbacks
- ⚠️ Macros must be enabled (security warning)
- ⚠️ Doesn't work in Excel Online (desktop only)

### Gameplay
- ⚠️ Flappy Bird collision could be more forgiving
- ⚠️ No sound effects (VBA `Beep` is too basic)
- ⚠️ No high score persistence across sessions

## 🚀 Future Improvements

### Potential Features
- 🔊 **Sound System** - Use Windows API for better audio
- 💾 **High Score Tracking** - Save to hidden sheet or external file
- 🎨 **More Pixel Art** - Additional enemy sprites
- 🎮 **Power-ups** - Speed boost, shield, multi-ball
- 🏆 **Achievement System** - Unlock skins, modes
- 👥 **Two-Player Mode** - Competitive Pong
- 🌈 **Visual Effects** - Screen shake, particle explosions
- 📱 **Touch Controls** - Better macro button placement

### Code Refactoring
- Separate rendering engine from game logic
- Implement proper game state machine
- Add configuration file for easy tuning
- Create reusable animation framework

## 📚 What I Learned

### Technical Skills
- ✅ VBA advanced techniques (timers, user-defined types, modules)
- ✅ Game loop architecture and timing
- ✅ Collision detection algorithms
- ✅ Animation and physics simulation
- ✅ Excel object model deep dive

### Game Design
- ✅ Balancing difficulty curves
- ✅ Player feedback systems (visual cues)
- ✅ Progressive challenge design
- ✅ Importance of playtesting

### Hackathon Lessons
- ✅ Scope management in time-limited projects
- ✅ Rapid prototyping and iteration
- ✅ Creative problem-solving with constraints
- ✅ Making unconventional choices that stand out

## 🙏 Acknowledgments

- **NUS Hackers** - For hosting an amazing hackathon
- **Classic Arcade Games** - Inspiration from Pong (1972) and Flappy Bird (2013)
- **Excel Community** - For VBA documentation and examples

---

**Created in 24 hours for Hack&Roll 2026**  
*Proving that Excel is not just for spreadsheets!* 📊➡️🎮
