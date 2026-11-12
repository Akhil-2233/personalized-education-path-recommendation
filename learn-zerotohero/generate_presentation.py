"""
Generate PowerPoint presentation for Personalized Learning Platform
Technical Architecture Focused
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_presentation():
    """Create a comprehensive technical presentation"""
    
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Define color scheme
    DARK_BLUE = RGBColor(0, 51, 102)
    ACCENT_BLUE = RGBColor(0, 102, 204)
    TEAL = RGBColor(0, 153, 153)
    LIGHT_GRAY = RGBColor(240, 240, 240)
    WHITE = RGBColor(255, 255, 255)
    TEXT_COLOR = RGBColor(33, 33, 33)
    
    def add_title_slide(title, subtitle):
        """Add a title slide"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = DARK_BLUE
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.text = title
        title_paragraph.font.size = Pt(54)
        title_paragraph.font.bold = True
        title_paragraph.font.color.rgb = WHITE
        title_paragraph.alignment = PP_ALIGN.CENTER
        
        # Subtitle
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(2))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.word_wrap = True
        subtitle_paragraph = subtitle_frame.paragraphs[0]
        subtitle_paragraph.text = subtitle
        subtitle_paragraph.font.size = Pt(28)
        subtitle_paragraph.font.color.rgb = ACCENT_BLUE
        subtitle_paragraph.alignment = PP_ALIGN.CENTER
        
        return slide
    
    def add_content_slide(title, content_items):
        """Add a content slide with bullet points"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = WHITE
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.text = title
        title_paragraph.font.size = Pt(44)
        title_paragraph.font.bold = True
        title_paragraph.font.color.rgb = DARK_BLUE
        
        # Title underline
        slide.shapes.add_shape(1, Inches(0.5), Inches(1.15), Inches(9), Inches(0)).line.color.rgb = ACCENT_BLUE
        
        # Content
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(5.5))
        text_frame = content_box.text_frame
        text_frame.word_wrap = True
        
        for i, item in enumerate(content_items):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = item
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_COLOR
            p.level = 0
            p.space_before = Pt(6)
            p.space_after = Pt(6)
        
        return slide
    
    def add_two_column_slide(title, left_items, right_items):
        """Add a slide with two columns"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = WHITE
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.text = title
        title_paragraph.font.size = Pt(44)
        title_paragraph.font.bold = True
        title_paragraph.font.color.rgb = DARK_BLUE
        
        # Left column
        left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(4.3), Inches(5.8))
        left_frame = left_box.text_frame
        left_frame.word_wrap = True
        
        for i, item in enumerate(left_items):
            if i == 0:
                p = left_frame.paragraphs[0]
            else:
                p = left_frame.add_paragraph()
            p.text = item
            p.font.size = Pt(14)
            p.font.color.rgb = TEXT_COLOR
            p.space_before = Pt(4)
            p.space_after = Pt(4)
        
        # Right column
        right_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.3), Inches(4.3), Inches(5.8))
        right_frame = right_box.text_frame
        right_frame.word_wrap = True
        
        for i, item in enumerate(right_items):
            if i == 0:
                p = right_frame.paragraphs[0]
            else:
                p = right_frame.add_paragraph()
            p.text = item
            p.font.size = Pt(14)
            p.font.color.rgb = TEXT_COLOR
            p.space_before = Pt(4)
            p.space_after = Pt(4)
        
        return slide
    
    # ============ SLIDE 1: TITLE SLIDE ============
    add_title_slide(
        "📚 Personalized Learning Platform",
        "Technical Architecture & Implementation\nB.Tech Final Year Project"
    )
    
    # ============ SLIDE 2: PROJECT OVERVIEW ============
    add_content_slide(
        "Project Overview",
        [
            "• Comprehensive Flask-based learning platform",
            "• Real-time emotion detection using webcam",
            "• 13+ interactive programming courses",
            "• 7 AI-powered career pathways",
            "• Advanced learning analytics dashboard",
            "• SQLite database with 8+ normalized tables",
            "• Multi-modal face detection (OpenCV, MediaPipe, dlib)"
        ]
    )
    
    # ============ SLIDE 3: TECHNOLOGY STACK ============
    add_two_column_slide(
        "Technology Stack",
        [
            "🐍 Backend:",
            "  • Python 3.8+",
            "  • Flask (web framework)",
            "  • SQLite (database)",
            "  • Werkzeug (security)",
            "",
            "🎨 Frontend:",
            "  • HTML5 / CSS3",
            "  • JavaScript (ES6+)",
            "  • Bootstrap (UI)",
            "  • Responsive Design"
        ],
        [
            "🤖 Machine Learning:",
            "  • OpenCV (image processing)",
            "  • TensorFlow (deep learning)",
            "  • PyTorch (neural networks)",
            "  • MediaPipe (face detection)",
            "  • dlib (facial landmarks)",
            "  • Scikit-learn (ML algorithms)",
            "",
            "📊 Data Processing:",
            "  • Pandas (data manipulation)",
            "  • Matplotlib (visualization)",
            "  • NumPy (numerical computing)"
        ]
    )
    
    # ============ SLIDE 4: ARCHITECTURE OVERVIEW ============
    add_content_slide(
        "System Architecture",
        [
            "🏗️ Three-Layer Architecture:",
            "",
            "1️⃣  Presentation Layer:",
            "   • 33 HTML templates with Bootstrap responsive design",
            "   • Real-time JavaScript for interactivity",
            "   • Webcam emotion detection interface",
            "",
            "2️⃣  Business Logic Layer:",
            "   • Flask application (app.py - 2583 lines)",
            "   • User authentication and session management",
            "   • Progress tracking and analytics",
            "",
            "3️⃣  Data Layer:",
            "   • SQLite database with 8+ normalized tables",
            "   • Query optimization with indexed foreign keys"
        ]
    )
    
    # ============ SLIDE 5: DATABASE DESIGN ============
    add_content_slide(
        "Database Architecture",
        [
            "📊 8 Core Tables with Normalized Design:",
            "",
            "👥 users → user profiles, authentication",
            "📚 courses → course metadata, description",
            "📈 progress → learning progress tracking",
            "🎯 sub_exercises → exercise content and metadata",
            "📝 sub_exercise_progress → user exercise tracking",
            "😊 emotions → emotion data logging",
            "🧪 quiz_results → quiz performance tracking",
            "📋 learning_actions → user activity logging",
            "",
            "🔑 Foreign Key Relationships for data integrity",
            "🔍 Indexed queries for fast retrieval"
        ]
    )
    
    # ============ SLIDE 6: CORE MODULES ============
    add_two_column_slide(
        "Core Application Modules",
        [
            "📄 app.py (2583 lines)",
            "  • Flask application initialization",
            "  • 50+ route endpoints",
            "  • Authentication (login/signup)",
            "  • Course serving",
            "  • Quiz handling",
            "  • Emotion detection API",
            "",
            "📦 db.py (765 lines)",
            "  • Database initialization",
            "  • CRUD operations",
            "  • User management",
            "  • Progress tracking queries"
        ],
        [
            "🤖 ml_models.py",
            "  • Emotion prediction",
            "  • Face detection",
            "  • Concentration scoring",
            "  • Multi-method detection",
            "",
            "📊 learning_analytics.py",
            "  • Analytics computation",
            "  • Performance metrics",
            "  • Data aggregation",
            "",
            "🛠️ utils.py",
            "  • Helper functions",
            "  • Data validation",
            "  • Formatting utilities"
        ]
    )
    
    # ============ SLIDE 7: FEATURES - LEARNING SYSTEM ============
    add_content_slide(
        "Core Features: Learning System",
        [
            "✅ Interactive Courses:",
            "   • 13+ courses covering programming, ML, cloud, web dev",
            "   • Structured modules with progression tracking",
            "   • 5 sub-exercise types: Theory, Examples, Practice, Quiz, Project",
            "",
            "✅ Quiz & Assessment:",
            "   • 20+ questions per difficulty level",
            "   • Auto-grading with instant feedback",
            "   • Detailed explanations for every answer",
            "   • Timed quiz sessions",
            "",
            "✅ Progress Tracking:",
            "   • Real-time dashboard visualization",
            "   • Completion percentage per course",
            "   • Time spent analytics",
            "   • Performance metrics"
        ]
    )
    
    # ============ SLIDE 8: FEATURES - EMOTION DETECTION ============
    add_content_slide(
        "Advanced Feature: Emotion Detection",
        [
            "🎥 Real-Time Emotion & Concentration Detection:",
            "",
            "🔬 Emotion Classification:",
            "   • Happy, Focused, Neutral, Tired, Distracted, Angry, Sad, etc.",
            "   • Concentration scoring (0-100%)",
            "   • Real-time feedback on engagement",
            "",
            "🧠 Multi-Method Face Detection:",
            "   • Primary: OpenCV Haar Cascades (fast)",
            "   • Secondary: MediaPipe (accurate, robust)",
            "   • Tertiary: dlib (advanced landmarks)",
            "",
            "📡 Integration:",
            "   • POST /api/emotion_detect endpoint",
            "   • Base64 image processing",
            "   • JSON response with emotion + concentration",
            "   • Emotion history logged to database"
        ]
    )
    
    # ============ SLIDE 9: AVAILABLE COURSES ============
    add_two_column_slide(
        "Available Courses (13+)",
        [
            "Frontend Development:",
            "  1. HTML & CSS Basics",
            "  2. JavaScript Essentials",
            "  3. React Development",
            "  4. Advanced CSS & Responsive",
            "",
            "Backend Development:",
            "  5. Java Programming",
            "  6. Node.js Backend",
            "  7. PHP Web Development",
            "",
            "Specialized:",
            "  8. Machine Learning",
            "  9. Data Structures & Algorithms"
        ],
        [
            "DevOps & Cloud:",
            "  10. Docker & Containerization",
            "  11. AWS Cloud Computing",
            "",
            "Difficulty Levels:",
            "  🟢 Basic (Fundamentals)",
            "  🟡 Intermediate (Core concepts)",
            "  🔴 Advanced (Expert level)",
            "",
            "Quiz Bank:",
            "  • 20+ questions per level",
            "  • 12+ JSON question files",
            "  • Auto-grading system"
        ]
    )
    
    # ============ SLIDE 10: CAREER PATHWAYS ============
    add_content_slide(
        "AI-Powered Career Pathways (7 Paths)",
        [
            "🧠 1. Machine Learning Developer",
            "   Skills: ML algorithms, Python, Data Structures",
            "",
            "⚙️  2. Backend Developer",
            "   Skills: Java, Node.js, PHP, DSA",
            "",
            "🎨 3. Frontend Developer",
            "   Skills: HTML/CSS, JavaScript, React",
            "",
            "🔗 4. Full Stack Developer",
            "   Skills: Complete web development stack",
            "",
            "⚡ 5. DevOps Engineer",
            "   Skills: Docker, AWS, Infrastructure",
            "",
            "🌐 6. Web Developer + 🎯 7. DSA Specialist"
        ]
    )
    
    # ============ SLIDE 11: USER AUTHENTICATION ============
    add_content_slide(
        "Security & Authentication",
        [
            "🔐 User Authentication System:",
            "   • Secure registration with email validation",
            "   • Password hashing using Werkzeug.security",
            "   • Bcrypt-equivalent security",
            "   • Session-based authentication",
            "",
            "👤 User Profiles:",
            "   • Customizable bio and learning goals",
            "   • Avatar upload support",
            "   • Progress persistence across sessions",
            "",
            "🔒 Security Features:",
            "   • Flask secret key configuration",
            "   • SQL injection prevention (parameterized queries)",
            "   • Input validation on client and server",
            "   • CSRF token support",
            "",
            "Default Credentials (Development):",
            "   Email: admin@example.com | Password: admin123"
        ]
    )
    
    # ============ SLIDE 12: LEARNING ANALYTICS ============
    add_two_column_slide(
        "Learning Analytics & Insights",
        [
            "📊 Dashboard Metrics:",
            "  • Total courses enrolled",
            "  • Overall progress percentage",
            "  • Time spent per course",
            "  • Average concentration score",
            "  • Emotion distribution analysis",
            "  • Quiz performance trends",
            "  • Completed modules count",
            "",
            "📈 Real-Time Tracking:",
            "  • Module completion tracking",
            "  • Score persistence",
            "  • Time analytics per exercise"
        ],
        [
            "💡 Individual Analytics:",
            "  • Progress per topic",
            "  • Exercise attempt history",
            "  • Quiz scores and retakes",
            "  • Emotional engagement levels",
            "  • Strong areas identification",
            "  • Weak areas needing focus",
            "",
            "🎯 AI Recommendations:",
            "  • Next course suggestions",
            "  • Difficulty progression",
            "  • Career path alignment",
            "  • Learning gap identification"
        ]
    )
    
    # ============ SLIDE 13: DASHBOARD UI/UX ============
    add_content_slide(
        "Dashboard & User Interface",
        [
            "🎨 Responsive Design Features:",
            "   • Mobile-first approach",
            "   • Bootstrap framework integration",
            "   • CSS3 animations and transitions",
            "   • Dark/Light theme support",
            "",
            "📱 Dashboard Components:",
            "   • Course enrollment cards",
            "   • Progress bars with percentage",
            "   • Learning statistics section",
            "   • Career pathways explorer",
            "   • Concentration detection widget",
            "   • Quiz result summaries",
            "",
            "🖼️ Template Architecture:",
            "   • 33 HTML templates",
            "   • Jinja2 template inheritance",
            "   • Dynamic content rendering",
            "   • Form validation feedback"
        ]
    )
    
    # ============ SLIDE 14: FRONT-END CODE ============
    add_content_slide(
        "Frontend Implementation",
        [
            "📁 Static Assets Structure:",
            "   • CSS/ (8+ stylesheets)",
            "     - dashboard.css, ui-enhanced.css, face-detection.css",
            "   • JS/ (6+ JavaScript files)",
            "     - webcam_hybrid.js (emotion detection)",
            "     - quiz.js (quiz interactions)",
            "     - face_concentration.js (concentration tracking)",
            "     - enhanced-exercises.js (exercise feedback)",
            "   • Models/ (ML model files)",
            "   • PDFs/ (course materials)",
            "",
            "🔌 API Integration:",
            "   • AJAX calls for real-time updates",
            "   • WebSocket support for live features",
            "   • Base64 image encoding for emotion detection"
        ]
    )
    
    # ============ SLIDE 15: DEPLOYMENT & CONFIGURATION ============
    add_content_slide(
        "Deployment & Configuration",
        [
            "🚀 Quick Start Options:",
            "",
            "Option 1: One-Click (Recommended)",
            "   • Double-click start.bat",
            "   • Auto: venv setup, dependencies, DB init, launch",
            "",
            "Option 2: Manual Setup",
            "   • python -m venv .venv",
            "   • .venv\\Scripts\\activate",
            "   • pip install -r requirements.txt",
            "   • python db.py",
            "   • python app.py",
            "",
            "⚙️ Default Configuration:",
            "   • Host: localhost (127.0.0.1)",
            "   • Port: 5000",
            "   • Debug: Enabled (development)",
            "   • Database: SQLite (app.db)"
        ]
    )
    
    # ============ SLIDE 16: PROJECT STATISTICS ============
    add_two_column_slide(
        "Project Statistics & Metrics",
        [
            "📊 Codebase Metrics:",
            "  • 2583 lines: app.py (main)",
            "  • 765 lines: db.py (database)",
            "  • 25+ dependencies installed",
            "  • 33 HTML templates",
            "  • 8+ CSS stylesheets",
            "  • 6+ JavaScript modules",
            "",
            "📚 Content Volume:",
            "  • 13+ courses",
            "  • 12+ quiz JSON files",
            "  • 20+ Q/A per level",
            "  • 7 career pathways"
        ],
        [
            "🗄️ Database Schema:",
            "  • 8 core tables",
            "  • 12+ relationships",
            "  • Multiple indexes",
            "  • Normalized design",
            "",
            "🔗 API Endpoints:",
            "  • 50+ Flask routes",
            "  • RESTful design",
            "  • JSON responses",
            "  • Real-time emotion detection endpoint",
            "",
            "⏱️ Performance:",
            "  • Sub-second page loads (optimized)",
            "  • Real-time emotion detection",
            "  • Efficient database queries"
        ]
    )
    
    # ============ SLIDE 17: SYSTEM REQUIREMENTS ============
    add_content_slide(
        "System Requirements & Compatibility",
        [
            "💻 Hardware Requirements:",
            "   • Python 3.8+ (3.12 recommended)",
            "   • RAM: 4GB minimum, 8GB recommended for ML models",
            "   • Disk: 2GB for dependencies and models",
            "   • Webcam: Optional (for emotion detection)",
            "",
            "🌏 Browser Compatibility:",
            "   ✅ Google Chrome (Latest)",
            "   ✅ Mozilla Firefox (Latest)",
            "   ✅ Microsoft Edge (Latest)",
            "   ✅ Safari (Latest)",
            "   ✅ Mobile Browsers (Responsive)",
            "",
            "🔧 Dependencies:",
            "   • Flask, OpenCV, TensorFlow, PyTorch, MediaPipe",
            "   • Complete stack in requirements.txt"
        ]
    )
    
    # ============ SLIDE 18: TESTING & VALIDATION ============
    add_content_slide(
        "Testing & Quality Assurance",
        [
            "✅ Feature Testing:",
            "   • User authentication validation",
            "   • Course enrollment and progress tracking",
            "   • Quiz auto-grading accuracy",
            "   • Emotion detection edge cases",
            "",
            "🐛 Quality Checks:",
            "   • Database integrity validation",
            "   • Input sanitization testing",
            "   • Response time performance",
            "   • Cross-browser compatibility",
            "",
            "🔍 Debugging Tools:",
            "   • Console logging (development)",
            "   • Browser DevTools support",
            "   • Application logs directory",
            "   • Error tracking and reporting"
        ]
    )
    
    # ============ SLIDE 19: CHALLENGES & SOLUTIONS ============
    add_two_column_slide(
        "Development Challenges & Solutions",
        [
            "⚠️ Challenges Addressed:",
            "",
            "1. Real-time Emotion Detection",
            "   ✓ Multi-method approach (3 algorithms)",
            "   ✓ Fallback mechanisms",
            "   ✓ Efficient image processing",
            "",
            "2. Database Performance",
            "   ✓ Normalized schema design",
            "   ✓ Query optimization",
            "   ✓ Foreign key indexing",
            "",
            "3. User Experience",
            "   ✓ Responsive design",
            "   ✓ Real-time updates",
            "   ✓ Fast load times"
        ],
        [
            "4. ML Model Integration",
            "   ✓ OpenCV + TensorFlow",
            "   ✓ Async processing",
            "   ✓ Hardware acceleration",
            "",
            "5. Scalability",
            "   ✓ Modular architecture",
            "   ✓ Separation of concerns",
            "   ✓ Easy to extend",
            "",
            "6. Security",
            "   ✓ Password hashing",
            "   ✓ Session management",
            "   ✓ Input validation"
        ]
    )
    
    # ============ SLIDE 20: LEARNING OUTCOMES ============
    add_content_slide(
        "Learning Outcomes for B.Tech Students",
        [
            "🎓 Technical Skills Acquired:",
            "",
            "Backend Development:",
            "   ✓ Flask framework and routing",
            "   ✓ Database design and optimization",
            "   ✓ User authentication and security",
            "   ✓ API development",
            "",
            "Frontend Development:",
            "   ✓ HTML5/CSS3 responsive design",
            "   ✓ JavaScript ES6+ programming",
            "   ✓ Real-time data updates",
            "   ✓ User interface design patterns",
            "",
            "Machine Learning & Computer Vision:",
            "   ✓ OpenCV image processing",
            "   ✓ Neural networks (TensorFlow, PyTorch)",
            "   ✓ Real-time ML inference",
            "   ✓ Emotion detection algorithms"
        ]
    )
    
    # ============ SLIDE 21: FUTURE ENHANCEMENTS ============
    add_two_column_slide(
        "Future Enhancement Roadmap",
        [
            "🎮 Gamification:",
            "  • Badges and achievements",
            "  • Leaderboards",
            "  • Scoring system",
            "  • Streak tracking",
            "",
            "🤖 AI Enhancements:",
            "  • Personalized recommendations",
            "  • Adaptive learning paths",
            "  • Predictive models",
            "",
            "📱 Mobile:",
            "  • Progressive Web App",
            "  • Native mobile app",
            "  • Offline support"
        ],
        [
            "🌍 Collaboration:",
            "  • Discussion forums",
            "  • Peer code reviews",
            "  • Study groups",
            "  • Mentorship matching",
            "",
            "📊 Monetization:",
            "  • Course marketplace",
            "  • Certification system",
            "  • Premium features",
            "  • Payment integration",
            "",
            "🔧 Infrastructure:",
            "  • Cloud deployment (AWS/GCP)",
            "  • Database scaling",
            "  • API rate limiting"
        ]
    )
    
    # ============ SLIDE 22: PROJECT IMPACT ============
    add_content_slide(
        "Educational Impact & Real-World Application",
        [
            "📈 Benefits to Students:",
            "   • Self-paced personalized learning",
            "   • Real-time engagement feedback via emotion detection",
            "   • Career pathway guidance based on skills",
            "   • Comprehensive progress analytics",
            "   • Interactive hands-on exercises",
            "",
            "👨‍🏫 Benefits to Educators:",
            "   • Automated grading system",
            "   • Student engagement metrics",
            "   • Learning pattern analytics",
            "   • Easy course content management",
            "",
            "🏢 Real-World Applicability:",
            "   • Scalable to enterprise learning management",
            "   • Integration with corporate training programs",
            "   • Government e-learning initiatives",
            "   • Online course platforms"
        ]
    )
    
    # ============ SLIDE 23: TECHNICAL EXCELLENCE ============
    add_content_slide(
        "Technical Excellence & Best Practices",
        [
            "✨ Software Engineering Principles:",
            "   • MVC architecture pattern",
            "   • Separation of concerns",
            "   • Database normalization",
            "   • RESTful API design",
            "   • Code modularity",
            "",
            "🛡️ Quality Standards:",
            "   • Input validation and sanitization",
            "   • Error handling and logging",
            "   • Performance optimization",
            "   • Security best practices",
            "   • Documentation",
            "",
            "🔄 Development Practices:",
            "   • Version control ready",
            "   • Modular code structure",
            "   • Configuration management",
            "   • Deployment automation"
        ]
    )
    
    # ============ SLIDE 24: CONCLUSION ============
    add_title_slide(
        "Thank You!",
        "Questions & Discussion\n\nPersonalized Learning Platform\nwww.github.com/your-repo"
    )
    
    # Save presentation
    output_path = "Personalized_Learning_Platform_Technical_Presentation.pptx"
    prs.save(output_path)
    print(f"✅ Presentation created successfully: {output_path}")
    return output_path

if __name__ == "__main__":
    create_presentation()
