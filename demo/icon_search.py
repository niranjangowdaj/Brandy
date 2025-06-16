import sys
import os
import re
from difflib import SequenceMatcher
from collections import defaultdict

class IconSearcher:
    def __init__(self, library_path):
        self.library_path = library_path
        self.icons = self._load_icons()
        
    def _load_icons(self):
        """Load all icons from the library and extract metadata"""
        icons = []
        
        if not os.path.exists(self.library_path):
            print(f"Error: Icon library path '{self.library_path}' not found.")
            return icons
            
        for filename in os.listdir(self.library_path):
            if filename.endswith(('.png', '.svg')):
                # Parse filename: ID_name_color.extension
                parts = filename.split('_')
                if len(parts) >= 3:
                    icon_id = parts[0]
                    # Handle cases where there might be duplicate IDs in name
                    name_parts = parts[1:-1]  # Everything except first (ID) and last (color.ext)
                    name = '_'.join(name_parts)
                    
                    # Extract color and extension
                    color_ext = parts[-1]
                    color = color_ext.split('.')[0]
                    extension = color_ext.split('.')[1]
                    
                    icons.append({
                        'filename': filename,
                        'id': icon_id,
                        'name': name,
                        'color': color,
                        'extension': extension,
                        'path': os.path.join(self.library_path, filename),
                        'keywords': self._extract_keywords(name)
                    })
        
        return icons
    
    def _extract_keywords(self, name):
        """Extract searchable keywords from icon name"""
        # Split by hyphens and underscores, convert to lowercase
        keywords = re.split(r'[-_]', name.lower())
        # Add the full name as well
        keywords.append(name.lower())
        # Remove empty strings
        keywords = [k for k in keywords if k]
        return keywords
    
    def _calculate_similarity(self, search_term, icon):
        """Calculate similarity score between search term and icon"""
        search_term = search_term.lower()
        max_score = 0
        
        # Check exact matches first (highest score)
        for keyword in icon['keywords']:
            if search_term == keyword:
                return 1.0
            
            # Check if search term is contained in keyword (must be meaningful)
            if search_term in keyword and len(search_term) >= 3:
                score = len(search_term) / len(keyword)
                max_score = max(max_score, score * 0.9)
            
            # Check if keyword is contained in search term (must be meaningful)
            if keyword in search_term and len(keyword) >= 3:
                score = len(keyword) / len(search_term)
                max_score = max(max_score, score * 0.8)
            
            # Use sequence matching for fuzzy matching (stricter threshold)
            similarity = SequenceMatcher(None, search_term, keyword).ratio()
            if similarity > 0.6:  # Only consider high similarity matches
                max_score = max(max_score, similarity * 0.7)
        
        return max_score
    
    def search(self, search_term, min_score=0.5, max_results=20):  # Increased min_score
        """Search for icons matching the search term"""
        results = []
        
        for icon in self.icons:
            score = self._calculate_similarity(search_term, icon)
            if score >= min_score:
                results.append((score, icon))
        
        # Sort by score (descending) and limit results
        results.sort(key=lambda x: x[0], reverse=True)
        return results[:max_results]
    
    def get_icon_variants(self, icon_name):
        """Get all color variants of a specific icon"""
        variants = []
        for icon in self.icons:
            if icon['name'] == icon_name:
                variants.append(icon)
        return variants
    
    def display_results(self, results, search_term):
        """Display search results in a formatted way"""
        if not results:
            print(f"❌ No relevant icons found for '{search_term}'")
            print()
            
            # Provide specific suggestions based on search term
            web3_terms = ['web3', 'blockchain', 'crypto', 'cryptocurrency', 'bitcoin', 'ethereum', 'nft', 'defi', 'dao', 'smart contract', 'metaverse']
            tech_terms = ['ai', 'artificial intelligence', 'machine learning', 'ml', 'iot', 'cloud computing', 'microservices']
            
            if search_term.lower() in web3_terms:
                print("🔗 Web3/Blockchain icons are not available in this library.")
                print("💡 Consider using these available alternatives:")
                print("   • 'network' - for blockchain/connectivity concepts")
                print("   • 'security' - for trust/verification concepts") 
                print("   • 'data' - for data/transaction concepts")
                print("   • 'computer' - for technology concepts")
                
            elif search_term.lower() in tech_terms:
                print("🤖 Advanced tech icons are limited in this library.")
                print("💡 Consider using these available alternatives:")
                print("   • 'data' - for AI/ML data concepts")
                print("   • 'network' - for connectivity/IoT concepts")
                print("   • 'computer' - for general technology")
                
            else:
                print("💡 Try using more general terms from available categories:")
                
                # Show top categories
                categories = defaultdict(int)
                for icon in self.icons:
                    for keyword in icon['keywords']:
                        if len(keyword) > 3:
                            categories[keyword] += 1
                
                top_categories = sorted(categories.items(), key=lambda x: x[1], reverse=True)[:10]
                for i, (category, count) in enumerate(top_categories, 1):
                    print(f"   {i:2d}. '{category}' ({count} icons)")
            
            print(f"\n🔍 Use 'python3 icon_search.py --categories' to see all {len(set(icon['name'] for icon in self.icons))} available icons")
            return
        
        print(f"\n🔍 Found {len(results)} relevant icons matching '{search_term}':")
        print("=" * 80)
        
        # Group by icon name to show variants together
        grouped = defaultdict(list)
        for score, icon in results:
            grouped[icon['name']].append((score, icon))
        
        for icon_name, variants in list(grouped.items())[:15]:  # Show top 15 icon groups
            # Sort variants by color preference (blue first, then white)
            variants.sort(key=lambda x: (x[1]['color'] != 'blue', x[1]['color'] != 'white'))
            
            best_score = max(score for score, _ in variants)
            print(f"\n📦 {icon_name.replace('-', ' ').replace('_', ' ').title()}")
            print(f"   Match Score: {best_score:.2f}")
            print(f"   Available variants:")
            
            for score, icon in variants:
                color_emoji = "🔵" if icon['color'] == 'blue' else "⚪" if icon['color'] == 'white' else "🟡"
                format_emoji = "🖼️" if icon['extension'] == 'png' else "📊" if icon['extension'] == 'svg' else "📄"
                print(f"     {color_emoji} {format_emoji} {icon['filename']}")
            
            print(f"   Path: {variants[0][1]['path'].replace(variants[0][1]['filename'], '')}")
    
    def suggest_categories(self):
        """Suggest icon categories based on available icons"""
        categories = defaultdict(list)
        
        for icon in self.icons:
            for keyword in icon['keywords']:
                if len(keyword) > 3:  # Ignore very short keywords
                    categories[keyword].append(icon['name'])
        
        # Sort by frequency
        sorted_categories = sorted(categories.items(), key=lambda x: len(x[1]), reverse=True)
        
        print("\n📂 Available icon categories (top 20):")
        print("-" * 50)
        for i, (category, icons) in enumerate(sorted_categories[:20], 1):
            print(f"{i:2d}. {category.title()} ({len(set(icons))} unique icons)")

def main():
    # Path to the icon library
    library_path = "ImageLibrary_60_20250609_1733"
    
    if len(sys.argv) < 2:
        print("Smart Icon Search Tool")
        print("Usage: python icon_search.py <search_term>")
        print("\nExamples:")
        print("  python icon_search.py security")
        print("  python icon_search.py network")
        print("  python icon_search.py data")
        print("  python icon_search.py --categories")
        sys.exit(1)
    
    searcher = IconSearcher(library_path)
    
    if not searcher.icons:
        print(f"No icons found in '{library_path}'. Please check the path.")
        sys.exit(1)
    
    search_term = sys.argv[1]
    
    if search_term == "--categories":
        searcher.suggest_categories()
    else:
        results = searcher.search(search_term)
        searcher.display_results(results, search_term)
        
        print(f"\n💡 Total icons in library: {len(searcher.icons)}")
        print("💡 Use --categories to see all available categories")

if __name__ == "__main__":
    main() 