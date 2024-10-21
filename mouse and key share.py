import tkinter as tk
from tkinter import messagebox
import asyncio
import threading
import json
from pynput import keyboard, mouse
import websockets
import os
import nest_asyncio
import screeninfo

class KMSharingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Keyboard and Mouse Sharing App")
        self.root.geometry("400x300")

        self.mode = tk.StringVar(value="server")
        self.server_ip = tk.StringVar()
        self.password = tk.StringVar()
        self.connection_status = tk.StringVar(value="Disconnected")

        # GUI Elements
        self.create_widgets()

        # WebSocket server/client variables
        self.clients = set()
        self.websocket = None
        self.server = None
        self.screen_width = screeninfo.get_monitors()[0].width

    def create_widgets(self):
        # Mode Selection
        tk.Label(self.root, text="Mode:").pack(pady=5)
        tk.Radiobutton(self.root, text="Server", variable=self.mode, value="server").pack()
        tk.Radiobutton(self.root, text="Client", variable=self.mode, value="client").pack()

        # Server IP Entry (for Client mode)
        tk.Label(self.root, text="Server IP:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.server_ip).pack()

        # Password Entry
        tk.Label(self.root, text="Password:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.password, show="*").pack()

        # Start/Stop Button
        self.start_button = tk.Button(self.root, text="Start", command=self.start)
        self.start_button.pack(pady=20)

        # Status Label
        tk.Label(self.root, text="Status:").pack()
        tk.Label(self.root, textvariable=self.connection_status).pack()

    def start(self):
        mode = self.mode.get()
        if mode == "server":
            self.start_server()
        elif mode == "client":
            if not self.server_ip.get():
                messagebox.showerror("Input Error", "Please enter the server IP.")
                return
            self.start_client()

    def start_server(self):
        self.connection_status.set("Starting server...")
        self.start_button.config(state=tk.DISABLED)

        # Start the server in a new thread
        threading.Thread(target=self.run_server, daemon=True).start()

    def run_server(self):
        nest_asyncio.apply()
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        async def handler(websocket, path):
            # Password verification
            try:
                auth_message = await websocket.recv()
                auth_data = json.loads(auth_message)
                if auth_data.get("password") != self.password.get():
                    await websocket.send(json.dumps({"error": "Authentication failed"}))
                    return
                self.clients.add(websocket)
                await websocket.send(json.dumps({"status": "Connected"}))
                async for message in websocket:
                    pass  # Server doesn't expect to receive messages from clients
            finally:
                self.clients.remove(websocket)

        async def broadcast(event_data):
            if self.clients:
                await asyncio.wait([client.send(event_data) for client in self.clients])

        def on_press(key):
            event = {'type': 'keyboard', 'action': 'press', 'key': str(key)}
            asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())

        def on_release(key):
            event = {'type': 'keyboard', 'action': 'release', 'key': str(key)}
            asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())

        def on_move(x, y):
            if x >= self.screen_width - 1:  # Move to client machine when mouse reaches edge of the screen
                event = {'type': 'mouse', 'action': 'move', 'position': (0, y)}
                asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())
            else:
                event = {'type': 'mouse', 'action': 'move', 'position': (x, y)}
                asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())

        def on_click(x, y, button, pressed):
            event = {'type': 'mouse', 'action': 'click', 'button': str(button), 'pressed': pressed, 'position': (x, y)}
            asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())

        def on_scroll(x, y, dx, dy):
            event = {'type': 'mouse', 'action': 'scroll', 'dx': dx, 'dy': dy, 'position': (x, y)}
            asyncio.run_coroutine_threadsafe(broadcast(json.dumps(event)), asyncio.get_event_loop())

        try:
            self.server = websockets.serve(handler, '0.0.0.0', 8765)
            
            # Start the keyboard and mouse listeners
            keyboard_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
            mouse_listener = mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll)
            keyboard_listener.start()
            mouse_listener.start()

            loop.run_until_complete(self.server)
            self.connection_status.set("Server running")
            loop.run_forever()
        except Exception as e:
            self.connection_status.set(f"Server error: {e}")
            self.start_button.config(state=tk.NORMAL)

    def start_client(self):
        self.connection_status.set("Starting client...")
        self.start_button.config(state=tk.DISABLED)

        # Start the client in a new thread
        threading.Thread(target=self.run_client, daemon=True).start()

    def run_client(self):
        nest_asyncio.apply()
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        uri = f"ws://{self.server_ip.get()}:8765"
        
        async def handle_events():
            try:
                async with websockets.connect(uri) as websocket:
                    # Send password for authentication
                    await websocket.send(json.dumps({"password": self.password.get()}))
                    response = await websocket.recv()
                    response_data = json.loads(response)
                    if "error" in response_data:
                        raise Exception(response_data["error"])
                    
                    self.connection_status.set("Client connected")
                    async for message in websocket:
                        event = json.loads(message)
                        if event['type'] == 'keyboard':
                            key = event['key']
                            if event['action'] == 'press':
                                keyboard.Controller().press(key)
                            elif event['action'] == 'release':
                                keyboard.Controller().release(key)
                        elif event['type'] == 'mouse':
                            if event['action'] == 'move':
                                x, y = event['position']
                                mouse.Controller().position = (x, y)
                            elif event['action'] == 'click':
                                button = mouse.Button[event['button'].lower()]
                                if event['pressed']:
                                    mouse.Controller().press(button)
                                else:
                                    mouse.Controller().release(button)
                            elif event['action'] == 'scroll':
                                dx = event['dx']
                                dy = event['dy']
                                mouse.Controller().scroll(dx, dy)
            except Exception as e:
                self.connection_status.set(f"Client error: {e}")
                self.start_button.config(state=tk.NORMAL)

        loop.run_until_complete(handle_events())

if __name__ == "__main__":
    root = tk.Tk()
    app = KMSharingApp(root)
    root.mainloop()

