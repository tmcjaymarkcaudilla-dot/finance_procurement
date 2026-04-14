<?php 
session_start(); // Siguraduhing may session_start() sa pinakataas
include "db.php"; 

// If user is already logged in, send them to the dashboard
if(isset($_SESSION['user'])){
    header("Location: dashboard.php");
    exit;
}

$error = "";

if(isset($_POST['login'])){
    $user = $conn->real_escape_string($_POST['username']);
    $pass = $_POST['password'];

    $result = $conn->query("SELECT * FROM users WHERE username='$user'");
    if($result && $result->num_rows > 0){
        $row = $result->fetch_assoc();
        if(password_verify($pass, $row['password'])){
            $_SESSION['user'] = $row['id'];
            header("Location: dashboard.php");
            exit;
        } else {
            $error = "Maling password! Subukan muli.";
        }
    } else {
        $error = "Hindi nahanap ang username!";
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - 16th IMB Finance System</title>
    <style>
        :root {
            --primary-green: #1b5e20;
            --hover-green: #165030;
            --accent-yellow: #fbc02d;
        }

        body {
            margin: 0;
            font-family: 'Segoe UI', Roboto, sans-serif;
            background: #f0f2f5;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }

        .login-card {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
            width: 100%;
            max-width: 380px;
            text-align: center;
            border-top: 6px solid var(--primary-green);
        }

        .login-card h2 {
            margin: 10px 0 5px;
            color: var(--primary-green);
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .system-subtitle {
            font-size: 13px;
            color: #666;
            margin-bottom: 25px;
            display: block;
        }

        .input-group {
            margin-bottom: 18px;
            text-align: left;
        }

        .input-group label {
            display: block;
            margin-bottom: 6px;
            color: #444;
            font-size: 14px;
            font-weight: 600;
        }

        .input-group input {
            width: 100%;
            padding: 12px;
            border: 1px solid #ccc;
            border-radius: 6px;
            box-sizing: border-box;
            font-size: 16px;
            transition: 0.3s;
        }

        .input-group input:focus {
            border-color: var(--primary-green);
            outline: none;
            box-shadow: 0 0 0 3px rgba(27, 94, 32, 0.1);
        }

        .login-btn {
            width: 100%;
            padding: 13px;
            background: var(--primary-green);
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: 0.3s;
            margin-top: 10px;
        }

        .login-btn:hover {
            background: var(--hover-green);
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }

        .register-box {
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            font-size: 14px;
            color: #555;
        }

        .register-link {
            color: var(--primary-green);
            text-decoration: none;
            font-weight: 700;
            transition: 0.2s;
        }

        .register-link:hover {
            text-decoration: underline;
            color: var(--hover-green);
        }

        .error-msg {
            color: #d32f2f;
            background: #ffebee;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
            border-left: 4px solid #d32f2f;
        }
    </style>
</head>
<body>

<div class="login-card">
    <h2>16th IMB PA</h2>
    <span class="system-subtitle">Finance Procurement Management System</span>

    <?php if($error): ?>
        <div class="error-msg"><?php echo $error; ?></div>
    <?php endif; ?>

    <form method="POST">
        <div class="input-group">
            <label>Username</label>
            <input type="text" name="username" placeholder="Gamitin ang iyong username" required>
        </div>
        
        <div class="input-group">
            <label>Password</label>
            <input type="password" name="password" placeholder="••••••••" required>
        </div>

        <button type="submit" name="login" class="login-btn">SIGN IN</button>
    </form>

    <div class="register-box">
        Wala pang account? 
        <a href="register.php" class="register-link">Register New Account</a>
    </div>
</div>

</body>
</html>