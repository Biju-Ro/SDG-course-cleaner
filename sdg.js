const fs = require('fs');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('sdgs.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(worksheet);

const sdgThemes = {
  'Social Foundation': [1, 2, 3, 4, 5, 10, 16],
  'Economic Growth': [8, 9, 17],
  'Environmental Protection': [6, 7, 11, 12, 13, 14, 15]
};

const sdgComplexity = {
  'Foundational': [1, 2, 3, 4, 5, 6],
  'Developmental': [7, 8, 9, 10, 11],
  'Transformational': [12, 13, 14, 15, 16, 17]
};

const sdgMap = {
  1: 'No poverty',
  2: 'Zero hunger', 
  3: 'Good health and well-being',
  4: 'Quality Education',
  5: 'Gender equality',
  6: 'Clean water and sanitation',
  7: 'Affordable and clean energy',
  8: 'Decent work and economic growth',
  9: 'Industry, innovation and infrastructure',
  10: 'Reduced inequalities',
  11: 'Sustainable cities and economies',
  12: 'Responsible consumption and production',
  13: 'Climate action',
  14: 'Life below water',
  15: 'Life on land',
  16: 'Peace, justice and strong institutions',
  17: 'Partnership for the goals'
};

function cleanValue(value) {
  if (!value || value.trim() === '' || value.trim() === ' ') return 'Unavailable';
  return value.trim();
}

function getSDGTheme(sdgNumber) {
  for (const [theme, sdgs] of Object.entries(sdgThemes)) {
    if (sdgs.includes(sdgNumber)) {
      return theme;
    }
  }
  return 'Other';
}

function getSDGComplexity(sdgNumber) {
  for (const [complexity, sdgs] of Object.entries(sdgComplexity)) {
    if (sdgs.includes(sdgNumber)) {
      return complexity;
    }
  }
  return 'Other';
}

function calculateSDGScope(sdgNumber) {
  const global = [1, 2, 13, 14, 15, 16, 17];
  const national = [3, 4, 5, 8, 9, 10];
  const local = [6, 7, 11, 12];
  
  if (global.includes(sdgNumber)) return 'Global';
  if (national.includes(sdgNumber)) return 'National';
  if (local.includes(sdgNumber)) return 'Local';
  return 'Multi-level';
}

function processData(data) {
  const normalizedData = [];
  const courseStats = new Map();
  const sdgStats = new Map();
  const themeStats = new Map();
  const complexityStats = new Map();
  const scopeStats = new Map();
  const professorStats = new Map();
  const uncertainSDGs = new Map();
  
  for (let i = 1; i <= 17; i++) {
    sdgStats.set(i, 0);
  }
  
  Object.keys(sdgThemes).forEach(theme => {
    themeStats.set(theme, 0);
  });
  
  Object.keys(sdgComplexity).forEach(complexity => {
    complexityStats.set(complexity, 0);
  });
  
  ['Global', 'National', 'Local', 'Multi-level'].forEach(scope => {
    scopeStats.set(scope, 0);
  });

  data.forEach(row => {
    const courseName = cleanValue(row['course name']);
    if (courseName === 'Unavailable') return;

    const professorEmail = cleanValue(row['Professor email']);
    const professorName = cleanValue(row['Professor name']);

    let courseSDGCount = 0;
    const courseThemes = new Set();
    const courseComplexities = new Set();
    const courseScopes = new Set();
    
    const uncertainSDG = cleanValue(row['SDG (not sure)']);
    if (uncertainSDG !== 'Unavailable') {
      if (!uncertainSDGs.has(uncertainSDG)) {
        uncertainSDGs.set(uncertainSDG, 0);
      }
      uncertainSDGs.set(uncertainSDG, uncertainSDGs.get(uncertainSDG) + 1);
    }

    for (let i = 1; i <= 17; i++) {
      const sdgValue = cleanValue(row[`SDG ${i}`]);
      if (sdgValue !== 'Unavailable') {
        courseSDGCount++;
        const theme = getSDGTheme(i);
        const complexity = getSDGComplexity(i);
        const scope = calculateSDGScope(i);
        
        courseThemes.add(theme);
        courseComplexities.add(complexity);
        courseScopes.add(scope);
        
        sdgStats.set(i, sdgStats.get(i) + 1);
        themeStats.set(theme, themeStats.get(theme) + 1);
        complexityStats.set(complexity, complexityStats.get(complexity) + 1);
        scopeStats.set(scope, scopeStats.get(scope) + 1);
        
        normalizedData.push({
          courseName: courseName,
          sdgNumber: i,
          sdgName: sdgValue,
          sdgStandardName: sdgMap[i],
          sdgTheme: theme,
          sdgComplexity: complexity,
          sdgScope: scope,
          
          courseSDGCount: 0,
          courseThemeCount: 0,
          courseComplexityCount: 0,
          courseScopeCount: 0,
          sdgFrequency: 0,
          themeFrequency: 0,
          complexityFrequency: 0,
          scopeFrequency: 0,
          
          hasUncertainSDG: uncertainSDG !== 'Unavailable' ? 1 : 0,
          uncertainSDGText: uncertainSDG,
          
          professorName: professorName,
          professorEmail: professorEmail
        });
      }
    }

    const sdgDensity = courseSDGCount / 17;
    const themeBalance = courseThemes.size / Object.keys(sdgThemes).length;
    const complexitySpread = courseComplexities.size / Object.keys(sdgComplexity).length;

    courseStats.set(courseName, {
      sdgCount: courseSDGCount,
      themeCount: courseThemes.size,
      complexityCount: courseComplexities.size,
      scopeCount: courseScopes.size,
      themes: Array.from(courseThemes),
      complexities: Array.from(courseComplexities),
      scopes: Array.from(courseScopes),
      sdgDensity: sdgDensity,
      themeBalance: themeBalance,
      complexitySpread: complexitySpread,
      hasUncertainSDG: uncertainSDG !== 'Unavailable',
      uncertainSDG: uncertainSDG,
      professor: {
        name: professorName,
        email: professorEmail
      }
    });

    if (professorEmail !== 'Unavailable') {
      if (!professorStats.has(professorEmail)) {
        professorStats.set(professorEmail, {
          name: professorName,
          courses: 0,
          totalSDGs: 0,
          uniqueSDGs: new Set(),
          uniqueThemes: new Set(),
          uniqueComplexities: new Set(),
          uniqueScopes: new Set(),
          avgSDGsPerCourse: 0,
          avgThemeBalance: 0,
          totalThemeBalance: 0,
          coursesWithUncertainSDGs: 0
        });
      }
      const prof = professorStats.get(professorEmail);
      prof.courses++;
      prof.totalSDGs += courseSDGCount;
      prof.totalThemeBalance += themeBalance;
      prof.avgSDGsPerCourse = prof.totalSDGs / prof.courses;
      prof.avgThemeBalance = prof.totalThemeBalance / prof.courses;
      
      if (uncertainSDG !== 'Unavailable') {
        prof.coursesWithUncertainSDGs++;
      }
      
      for (let i = 1; i <= 17; i++) {
        const sdgValue = cleanValue(row[`SDG ${i}`]);
        if (sdgValue !== 'Unavailable') {
          prof.uniqueSDGs.add(i);
          prof.uniqueThemes.add(getSDGTheme(i));
          prof.uniqueComplexities.add(getSDGComplexity(i));
          prof.uniqueScopes.add(calculateSDGScope(i));
        }
      }
    }
  });

  normalizedData.forEach(row => {
    const courseInfo = courseStats.get(row.courseName);
    row.courseSDGCount = courseInfo.sdgCount;
    row.courseThemeCount = courseInfo.themeCount;
    row.courseComplexityCount = courseInfo.complexityCount;
    row.courseScopeCount = courseInfo.scopeCount;
    row.sdgFrequency = sdgStats.get(row.sdgNumber);
    row.themeFrequency = themeStats.get(row.sdgTheme);
    row.complexityFrequency = complexityStats.get(row.sdgComplexity);
    row.scopeFrequency = scopeStats.get(row.sdgScope);
  });

  return { normalizedData, courseStats, sdgStats, themeStats, complexityStats, scopeStats, professorStats, uncertainSDGs };
}

function createSummaryData(normalizedData, courseStats, sdgStats, themeStats, complexityStats, scopeStats, professorStats, uncertainSDGs) {
  const totalCourses = courseStats.size;
  const totalMappings = normalizedData.length;
  
  const sdgFrequencyData = [];
  for (let i = 1; i <= 17; i++) {
    sdgFrequencyData.push({
      sdgNumber: i,
      sdgName: sdgMap[i],
      theme: getSDGTheme(i),
      complexity: getSDGComplexity(i),
      scope: calculateSDGScope(i),
      frequency: sdgStats.get(i),
      percentageOfCourses: ((sdgStats.get(i) / totalCourses) * 100).toFixed(2)
    });
  }
  sdgFrequencyData.sort((a, b) => b.frequency - a.frequency);

  const themeAnalysisData = [];
  themeStats.forEach((frequency, theme) => {
    const sdgsInTheme = sdgThemes[theme] || [];
    themeAnalysisData.push({
      theme: theme,
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2),
      sdgsInTheme: sdgsInTheme.join(', '),
      averageSDGsPerTheme: sdgsInTheme.length > 0 ? (frequency / sdgsInTheme.length).toFixed(2) : 0
    });
  });
  themeAnalysisData.sort((a, b) => b.frequency - a.frequency);

  const complexityAnalysisData = [];
  complexityStats.forEach((frequency, complexity) => {
    const sdgsInComplexity = sdgComplexity[complexity] || [];
    complexityAnalysisData.push({
      complexity: complexity,
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2),
      sdgsInComplexity: sdgsInComplexity.join(', ')
    });
  });
  complexityAnalysisData.sort((a, b) => b.frequency - a.frequency);

  const scopeAnalysisData = [];
  scopeStats.forEach((frequency, scope) => {
    scopeAnalysisData.push({
      scope: scope,
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
    });
  });
  scopeAnalysisData.sort((a, b) => b.frequency - a.frequency);

  const uncertainSDGData = [];
  uncertainSDGs.forEach((frequency, sdgText) => {
    uncertainSDGData.push({
      uncertainSDG: sdgText,
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
    });
  });
  uncertainSDGData.sort((a, b) => b.frequency - a.frequency);

  const courseComplexityData = [];
  courseStats.forEach((stats, courseName) => {
    courseComplexityData.push({
      courseName,
      sdgCount: stats.sdgCount,
      themeCount: stats.themeCount,
      complexityCount: stats.complexityCount,
      scopeCount: stats.scopeCount,
      sdgDensity: parseFloat((stats.sdgDensity * 100).toFixed(2)),
      themeBalance: parseFloat((stats.themeBalance * 100).toFixed(2)),
      complexitySpread: parseFloat((stats.complexitySpread * 100).toFixed(2)),
      themes: stats.themes.join(', '),
      complexities: stats.complexities.join(', '),
      scopes: stats.scopes.join(', '),
      hasUncertainSDG: stats.hasUncertainSDG ? 'Yes' : 'No',
      uncertainSDG: stats.uncertainSDG,
      overallRating: stats.sdgCount > 10 ? 'Comprehensive' : stats.sdgCount > 6 ? 'Extensive' : stats.sdgCount > 3 ? 'Moderate' : stats.sdgCount > 0 ? 'Limited' : 'None',
      professorName: stats.professor.name,
      professorEmail: stats.professor.email
    });
  });
  courseComplexityData.sort((a, b) => b.sdgCount - a.sdgCount);

  const professorAnalysisData = [];
  professorStats.forEach((stats, email) => {
    professorAnalysisData.push({
      professorEmail: email,
      professorName: stats.name,
      courseCount: stats.courses,
      totalSDGs: stats.totalSDGs,
      uniqueSDGs: stats.uniqueSDGs.size,
      uniqueThemes: stats.uniqueThemes.size,
      uniqueComplexities: stats.uniqueComplexities.size,
      uniqueScopes: stats.uniqueScopes.size,
      avgSDGsPerCourse: parseFloat(stats.avgSDGsPerCourse.toFixed(2)),
      avgThemeBalance: parseFloat((stats.avgThemeBalance * 100).toFixed(2)),
      coursesWithUncertainSDGs: stats.coursesWithUncertainSDGs,
      uncertaintyRate: parseFloat(((stats.coursesWithUncertainSDGs / stats.courses) * 100).toFixed(2)),
      sdgCoverage: parseFloat(((stats.uniqueSDGs.size / 17) * 100).toFixed(2)),
      themeCoverage: parseFloat(((stats.uniqueThemes.size / Object.keys(sdgThemes).length) * 100).toFixed(2))
    });
  });
  professorAnalysisData.sort((a, b) => b.sdgCoverage - a.sdgCoverage);

  const sdgDistribution = {
    coursesWithNoSDGs: totalCourses - new Set(normalizedData.map(row => row.courseName)).size,
    coursesWithSDGs: new Set(normalizedData.map(row => row.courseName)).size,
    coursesWithUncertainSDGs: Array.from(courseStats.values()).filter(course => course.hasUncertainSDG).length,
    averageSDGsPerCourse: (totalMappings / new Set(normalizedData.map(row => row.courseName)).size).toFixed(2),
    totalCourses,
    totalMappings,
    mostPopularSDG: sdgFrequencyData[0],
    leastPopularSDG: sdgFrequencyData[sdgFrequencyData.length - 1],
    mostPopularTheme: themeAnalysisData[0],
    mostCommonUncertainSDG: uncertainSDGData.length > 0 ? uncertainSDGData[0] : null
  };

  return { sdgFrequencyData, themeAnalysisData, complexityAnalysisData, scopeAnalysisData, uncertainSDGData, courseComplexityData, professorAnalysisData, sdgDistribution };
}

function exportToExcel(normalizedData, summaryData) {
  const workbook = XLSX.utils.book_new();

  const mainSheet = XLSX.utils.json_to_sheet(normalizedData.map(row => ({
    'Course Name': row.courseName,
    'SDG Number': row.sdgNumber,
    'SDG Standard Name': row.sdgStandardName,
    'SDG Theme': row.sdgTheme,
    'SDG Complexity': row.sdgComplexity,
    'SDG Scope': row.sdgScope,
    'Course SDG Count': row.courseSDGCount,
    'Course Theme Count': row.courseThemeCount,
    'Course Complexity Count': row.courseComplexityCount,
    'Course Scope Count': row.courseScopeCount,
    'SDG Frequency': row.sdgFrequency,
    'Theme Frequency': row.themeFrequency,
    'Complexity Frequency': row.complexityFrequency,
    'Scope Frequency': row.scopeFrequency,
    'Has Uncertain SDG': row.hasUncertainSDG,
    'Uncertain SDG Text': row.uncertainSDGText,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const sdgAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.sdgFrequencyData.map(row => ({
    'SDG Number': row.sdgNumber,
    'SDG Name': row.sdgName,
    'Theme': row.theme,
    'Complexity': row.complexity,
    'Scope': row.scope,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses
  })));

  const themeAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.themeAnalysisData.map(row => ({
    'Theme': row.theme,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses,
    'SDGs in Theme': row.sdgsInTheme,
    'Average SDGs per Theme': row.averageSDGsPerTheme
  })));

  const complexityAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.complexityAnalysisData.map(row => ({
    'Complexity Level': row.complexity,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses,
    'SDGs in Complexity': row.sdgsInComplexity
  })));

  const scopeAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.scopeAnalysisData.map(row => ({
    'Scope Level': row.scope,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses
  })));

  const uncertainSDGSheet = XLSX.utils.json_to_sheet(summaryData.uncertainSDGData.map(row => ({
    'Uncertain SDG': row.uncertainSDG,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses
  })));

  const courseAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.courseComplexityData.map(row => ({
    'Course Name': row.courseName,
    'SDG Count': row.sdgCount,
    'Theme Count': row.themeCount,
    'Complexity Count': row.complexityCount,
    'Scope Count': row.scopeCount,
    'SDG Density %': row.sdgDensity,
    'Theme Balance %': row.themeBalance,
    'Complexity Spread %': row.complexitySpread,
    'Themes': row.themes,
    'Complexities': row.complexities,
    'Scopes': row.scopes,
    'Has Uncertain SDG': row.hasUncertainSDG,
    'Uncertain SDG': row.uncertainSDG,
    'Overall Rating': row.overallRating,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const professorAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.professorAnalysisData.map(row => ({
    'Professor Email': row.professorEmail,
    'Professor Name': row.professorName,
    'Course Count': row.courseCount,
    'Total SDGs': row.totalSDGs,
    'Unique SDGs': row.uniqueSDGs,
    'Unique Themes': row.uniqueThemes,
    'Unique Complexities': row.uniqueComplexities,
    'Unique Scopes': row.uniqueScopes,
    'Avg SDGs per Course': row.avgSDGsPerCourse,
    'Avg Theme Balance %': row.avgThemeBalance,
    'Courses with Uncertain SDGs': row.coursesWithUncertainSDGs,
    'Uncertainty Rate %': row.uncertaintyRate,
    'SDG Coverage %': row.sdgCoverage,
    'Theme Coverage %': row.themeCoverage
  })));

  const overviewSheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total Courses', Value: summaryData.sdgDistribution.totalCourses },
    { Metric: 'Courses with SDGs', Value: summaryData.sdgDistribution.coursesWithSDGs },
    { Metric: 'Courses without SDGs', Value: summaryData.sdgDistribution.coursesWithNoSDGs },
    { Metric: 'Courses with Uncertain SDGs', Value: summaryData.sdgDistribution.coursesWithUncertainSDGs },
    { Metric: 'Total Course-SDG Mappings', Value: summaryData.sdgDistribution.totalMappings },
    { Metric: 'Average SDGs per Course', Value: summaryData.sdgDistribution.averageSDGsPerCourse },
    { Metric: 'Most Popular SDG', Value: `SDG ${summaryData.sdgDistribution.mostPopularSDG.sdgNumber}: ${summaryData.sdgDistribution.mostPopularSDG.sdgName}` },
    { Metric: 'Most Popular SDG Frequency', Value: summaryData.sdgDistribution.mostPopularSDG.frequency },
    { Metric: 'Least Popular SDG', Value: `SDG ${summaryData.sdgDistribution.leastPopularSDG.sdgNumber}: ${summaryData.sdgDistribution.leastPopularSDG.sdgName}` },
    { Metric: 'Least Popular SDG Frequency', Value: summaryData.sdgDistribution.leastPopularSDG.frequency },
    { Metric: 'Most Popular Theme', Value: summaryData.sdgDistribution.mostPopularTheme.theme },
    { Metric: 'Most Popular Theme Frequency', Value: summaryData.sdgDistribution.mostPopularTheme.frequency },
    { Metric: 'Most Common Uncertain SDG', Value: summaryData.sdgDistribution.mostCommonUncertainSDG ? summaryData.sdgDistribution.mostCommonUncertainSDG.uncertainSDG : 'None' }
  ]);

  XLSX.utils.book_append_sheet(workbook, mainSheet, 'Course-SDG Data');
  XLSX.utils.book_append_sheet(workbook, sdgAnalysisSheet, 'SDG Analysis');
  XLSX.utils.book_append_sheet(workbook, themeAnalysisSheet, 'Theme Analysis');
  XLSX.utils.book_append_sheet(workbook, complexityAnalysisSheet, 'Complexity Analysis');
  XLSX.utils.book_append_sheet(workbook, scopeAnalysisSheet, 'Scope Analysis');
  XLSX.utils.book_append_sheet(workbook, uncertainSDGSheet, 'Uncertain SDGs');
  XLSX.utils.book_append_sheet(workbook, courseAnalysisSheet, 'Course Analysis');
  XLSX.utils.book_append_sheet(workbook, professorAnalysisSheet, 'Professor Analysis');
  XLSX.utils.book_append_sheet(workbook, overviewSheet, 'Overview');

  XLSX.writeFile(workbook, 'comprehensive_sdgs_analysis.xlsx');
}

try {
  console.log('Processing comprehensive SDGs data for Excel analysis...');
  
  const { normalizedData, courseStats, sdgStats, themeStats, complexityStats, scopeStats, professorStats, uncertainSDGs } = processData(rawData);
  const summaryData = createSummaryData(normalizedData, courseStats, sdgStats, themeStats, complexityStats, scopeStats, professorStats, uncertainSDGs);
  
  exportToExcel(normalizedData, summaryData);
  
  console.log('\n📊 Comprehensive SDGs Analysis Complete!');
  console.log('📁 File created: comprehensive_sdgs_analysis.xlsx');
  console.log('\n📋 Excel sheets included:');
  console.log('   1. Course-SDG Data - Main normalized data with comprehensive metrics');
  console.log('   2. SDG Analysis - Individual SDG analysis with themes, complexity, scope');
  console.log('   3. Theme Analysis - Analysis by SDG thematic groupings');
  console.log('   4. Complexity Analysis - Analysis by SDG complexity levels');
  console.log('   5. Scope Analysis - Analysis by SDG scope (Global/National/Local)');
  console.log('   6. Uncertain SDGs - Analysis of uncertain/unclear SDG mappings');
  console.log('   7. Course Analysis - Comprehensive course metrics and ratings');
  console.log('   8. Professor Analysis - Detailed professor performance analysis');
  console.log('   9. Overview - Complete summary statistics');
  
  console.log('\n🔢 Comprehensive numeric columns for analysis:');
  console.log('   • Course SDG Count - Total SDGs per course');
  console.log('   • Course Theme Count - Number of themes covered per course');
  console.log('   • Course Complexity Count - Number of complexity levels per course');
  console.log('   • Course Scope Count - Number of scope levels per course');
  console.log('   • SDG Density % - Percentage of all 17 SDGs covered');
  console.log('   • Theme Balance % - Balance across thematic areas');
  console.log('   • Complexity Spread % - Distribution across complexity levels');
  
  console.log('\n🎯 SDG Classifications:');
  console.log('   Themes:');
  console.log('     • Social Foundation: SDGs 1, 2, 3, 4, 5, 10, 16');
  console.log('     • Economic Growth: SDGs 8, 9, 17');
  console.log('     • Environmental Protection: SDGs 6, 7, 11, 12, 13, 14, 15');
  console.log('   Complexity:');
  console.log('     • Foundational: SDGs 1, 2, 3, 4, 5, 6');
  console.log('     • Developmental: SDGs 7, 8, 9, 10, 11');
  console.log('     • Transformational: SDGs 12, 13, 14, 15, 16, 17');
  console.log('   Scope:');
  console.log('     • Global: SDGs 1, 2, 13, 14, 15, 16, 17');
  console.log('     • National: SDGs 3, 4, 5, 8, 9, 10');
  console.log('     • Local: SDGs 6, 7, 11, 12');
  
  console.log(`\n📈 Comprehensive Quick Stats:`);
  console.log(`   • Total courses: ${summaryData.sdgDistribution.totalCourses}`);
  console.log(`   • Total course-SDG pairs: ${summaryData.sdgDistribution.totalMappings}`);
  console.log(`   • Courses with uncertain SDGs: ${summaryData.sdgDistribution.coursesWithUncertainSDGs}`);
  console.log(`   • Average SDGs per course: ${summaryData.sdgDistribution.averageSDGsPerCourse}`);
  console.log(`   • Most popular SDG: SDG ${summaryData.sdgDistribution.mostPopularSDG.sdgNumber} (${summaryData.sdgDistribution.mostPopularSDG.frequency} courses)`);
  console.log(`   • Most popular theme: ${summaryData.sdgDistribution.mostPopularTheme.theme} (${summaryData.sdgDistribution.mostPopularTheme.frequency} instances)`);

} catch (error) {
  console.error('❌ Error processing data:', error);
}