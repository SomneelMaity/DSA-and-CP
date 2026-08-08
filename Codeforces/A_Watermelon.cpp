#include <bits/stdc++.h>
using namespace std;

// --- Optimizations ---
#pragma GCC optimize("O3,unroll-loops")
#pragma GCC target("avx2,bmi,bmi2,lzcnt,popcnt")

// --- Type Aliases ---
using ll = long long;
using ld = long double;
using pii = pair<int, int>;
using pll = pair<ll, ll>;
using vi = vector<int>;
using vll = vector<ll>;

// --- Macros ---
#define rep(i, a, b) for(int i = a; i < (b); ++i)
#define all(x) begin(x), end(x)
#define sz(x) (int)(x).size()
#define pb push_back
#define fi first
#define se second

// --- Local Debugging ---
#ifdef LOCAL
#include "debug.h" // Replace with your own debug header if you have one
#else
#define debug(...) 42
#endif

// --- Core Logic ---
void solve() {
    // Write your problem solution here
    int n = 0;
    cin>> n;
    if(n%2==0 && n>2){
        cout<< "YES\n";
    }
    else{
        cout<< "NO\n";
    }
}

int main() {
    // Optimize standard I/O operations for speed
    ios_base::sync_with_stdio(false);
    cin.tie(nullptr);
    
    int t = 1;
    // cin >> t; // Comment this out if the problem has only 1 test case
    while (t--) {
        solve();
    }
    
    return 0;
}
