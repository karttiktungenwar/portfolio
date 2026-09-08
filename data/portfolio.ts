// Central content file — edit this to update the site's text.
// Keeping content data-driven means the components stay generic
// and you never have to touch JSX just to change a word or a date.

export const profile = {
  name: "Kartik Tungenwar",
  role: "Senior Mobile App Developer & Designer(Android | Kotlin Multiplatform | Flutter)",
  tagline: "Building mobile apps with secure, production-grade release discipline.",
  location: "Secunderabad, Hyderabad, Telangana",
  email: "tungenwarkartick@gmail.com",
  phone: "+91 93225 50904",
  cv: "https://drive.google.com/file/d/1nQ6xlBToNKaglU-ITiW02f21RPHSLK2e/view?usp=sharing",
  social: {
    github: "https://github.com/karttiktungenwar",
    linkedin: "https://www.linkedin.com/in/kartik-tungenwar-b58867a0/",
    medium: "https://medium.com/@tungenwarkartick",
    whatsapp:"https://api.whatsapp.com/send/?phone=%2B919322550904&text&type=phone_number&app_absent=0",
    facebook: "https://www.facebook.com/profile.php?id=61557722770692",
    instagram: "https://www.instagram.com/rockstar_kartik_007/",
    microsoftteams: "https://teams.live.com/l/invite/FEA63xMS9Axou5lYwE?v=g1"
  }
};

export const about = {
  summary:
    "Seasoned Android developer with a demonstrated history in information technology and services, working across the full mobile app lifecycle — design, development, and production release.",
  detail:
    "Proficient in Java and Kotlin with a strong command over Android SDK, Jetpack, and Flutter. Comfortable with MVVM and similar modern architectures, RESTful API integration, secure data transmission, and building UI that stays out of the user's way."
};

export const skills = [
  { name: "Core Java", level: 100 },
  { name: "Android SDK", level: 90 },
  { name: "Kotlin / Jetpack / Flutter", level: 75 },
  { name: "SQL Query", level: 80 },
  { name: "React & Node JS", level: 55 },
  { name: "HTML, CSS & JavaScript", level: 55 }
];

export const experience = [
  {
    role: "Programming Analyst",
    company: "Mobicule Technologies Pvt. Ltd.",
    period: "Feb 2021 – Present",
    version: "v4.0.0",
    status: "current",
    notes: []
  },
  {
    role: "Senior Android Developer",
    company: "Clintico Infotech Pvt Ltd",
    period: "Jan 2019 – Dec 2020",
    version: "v3.0.0",
    status: "stable",
    notes: []
  },
  {
    role: "Junior Android Developer",
    company: "Cyborg Systems Nebula Studios",
    period: "Jan 2018 – Oct 2018",
    version: "v2.1.0",
    status: "stable",
    notes: []
  },
  {
    role: "Junior Android Developer",
    company: "Techflux Solutions",
    period: "May 2017 – Jan 2018",
    version: "v2.0.0",
    status: "stable",
    notes: []
  },
  {
    role: "Intern Android Developer",
    company: "Samkenishar Technoresearch",
    period: "Oct 2016 – Mar 2017",
    version: "v1.0.0",
    status: "archived",
    notes: []
  }
];

export const education = [
  {
    degree: "BE (E&TC)",
    school: "Savitribai Phule Pune University",
    period: "2009 – 2015"
  },
  {
    degree: "HSC (11th & 12th Science)",
    school: "Latur Divisional Board",
    period: "2007 – 2009"
  },
  {
    degree: "SSC (10th Grade)",
    school: "Latur Divisional Board",
    period: "2006 – 2007"
  }
];

export const projects = [
   {
    name: "My Lakshya",
    company: "Mobicule Technologies Pvt. Ltd.",
    period: "Mar 2025 – Present",
    version: "v3.4.0",
    tag: "stable",
    description:
      "Android enterprise app for telecom customer onboarding and SIM activation — digital eKYC, verification, document capture, authentication, postpaid activation, SIM exchange and migration workflows.",
    changes: [
      "+ Android app development and API integration",
      "+ UI enhancement and performance optimization",
      "+ SSL pinning security"
    ],
    link: "https://play.google.com/store/apps/details?id=com.mobicule.neeti.lakshyat&hl=en_IN"
  },
  {
    name: "vConnect",
    company: "Mobicule Technologies Pvt. Ltd.",
    period: "Mar 2023 – Present",
    version: "v3.4.0",
    tag: "stable",
    description:
      "Built a real-time workforce management and governance app used by Tata Play Fiber’s field teamsto optimize daily sales workflows and compliance.",
    changes: [
      "+ Flutter app development and API integration",
      "+ UI enhancement and performance optimization",
      "+ SSL pinning security, dynamic UI, Google Maps API"
    ],
    link: "https://play.google.com/store/apps/details?id=com.mobicule.vodafone.ekyc.client&hl=en_IN"
  },
  {
    name: "Tata Play – Disha",
    company: "Mobicule Technologies Pvt. Ltd.",
    period: "Mar 2023 – Present",
    version: "v2.8.0",
    tag: "stable",
    description:
      "Android enterprise app for field sales and operations — KPI tracking, visit management, performance monitoring, task execution, and real-time reporting.",
    changes: [
      "+ Android app development and API integration",
      "+ UI enhancement and performance optimization",
      "+ SSL pinning security, dynamic UI, Google Maps API"
    ],
    link: "https://play.google.com/store/apps/details?id=com.mobicule.tatasky&hl=en_IN"
  },
  {
    name: "DigiSales",
    company: "Mobicule Technologies Pvt. Ltd.",
    period: "Mar 2021 – Present",
    version: "v2.5.0",
    tag: "stable",
    description:
      "One-stop onboarding app for the Tata Play Fiber sales team.",
    changes: [
      "+ Android app development and API integration",
      "+ UI enhancement and performance optimization",
      "+ SSL pinning security, dynamic UI, Google Maps API"
    ],
    link: "https://play.google.com/store/apps/details?id=com.mobicule.tatasky_bb"
  },
  {
    name: "Agent Adda",
    company: "Clintico Infotech Pvt Ltd",
    period: "Nov 2020 – Dec 2020",
    version: "v1.2.0",
    tag: "stable",
    description:
      "App built for insurance agents to build business through digital marketing tools.",
    changes: [
      "+ Video and image watermarking",
      "+ Razorpay payment gateway integration"
    ],
    link: "https://play.google.com/store/apps/details?id=com.agentaddaapp"
  },
  {
    name: "Danga",
    company: "Clintico Infotech Pvt Ltd",
    period: "May 2020 – Oct 2022",
    version: "v1.6.0",
    tag: "stable",
    description:
      "Video creation and sharing app with personal chat, likes, follows and comments.",
    changes: [],
    link: "https://play.google.com/store/apps/details?id=com.danga&hl=en_IN"
  },
  {
    name: "Wishstore",
    company: "Cyborg Systems Nebula Studios",
    period: "Feb 2018 – Oct 2018",
    version: "v1.0.0",
    tag: "archived",
    description:
      "Retail kiosk platform helping shoppers find products at local retailers.",
    changes: [],
    link: "https://play.google.com/store/apps/details?id=com.csns.wishstorekioskcustomer"
  }
];

export const testimonials = [
  {
    quote:
      "The Mobile App Development program gave me the best post-secondary experience I could have acquired — the skills, knowledge and real-world experience needed for a career as a software developer.",
    name: "Myself",
    title: "Mobile App Developer"
  },
  {
    quote:
      "The Mobile Applications program is very well-rounded and teaches everything needed to become a successful, confident developer, ready for the workforce.",
    name: "Myself",
    title: "Mobile App Developer"
  },
  {
    quote:
      "Even with experience in the field, the Mobile Application Development program helped strengthen and update my skills with the latest technology.",
    name: "Myself",
    title: "Mobile App Developer"
  }
];
